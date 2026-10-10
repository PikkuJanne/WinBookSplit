"""Authored structural/cleanup controls; these are not native host passes."""

from copy import deepcopy
from hashlib import sha256
import importlib.util
import ast
import json
from pathlib import Path
import sys
import tempfile
import unittest
from types import SimpleNamespace
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[2]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


validator = load("wbs_cross_host_unit_guard", ROOT / "tests/faults/validate_cross_host_report.py")
fixture = load("wbs_cross_host_unit_console", ROOT / "tests/python/console_receipt_fixture.py")


def file_record(path, *, hashed="a", inode=1):
    return {"path": path, "exists": True, "device": 1, "inode": inode, "attributes": 32,
        "size_bytes": 100, "sha256": hashed * 64}


def host_record(host="PS51"):
    policy = [{"scope": "CurrentUser", "policy": "Undefined"}]
    return {"id": host, "passed": True, "shell_executable": "C:/" + host + "/" + ("powershell.exe" if host == "PS51" else "pwsh.exe"),
        "host_version": "5.1.26100.9444" if host == "PS51" else "7.6.6", "host_major": 5 if host == "PS51" else 7,
        "stored_policies": policy, "policies_after": deepcopy(policy), "exit_code": 0}


def complete_case(host="PS51", kind="replace"):
    observation = host_record(host)
    root = "C:/synthetic/" + host + "-captured-" + kind
    base, app, cwd, books = [root + "/" + name for name in ("o", "a", "c", "b")]
    source = books + "/processing.pdf"
    before = file_record(source)
    entries = [{"sequence": number, "title": f"Section (Page {start + 1}-{end})", "start": start, "end": end,
        "filename": f"{number:02} - Section (Page {start + 1}-{end}).pdf", "parent_id": None, "reason": "manual", "warnings": []}
        for number, (start, end) in enumerate(((0, 2), (2, 4)), 1)]
    source_id = {"path": source, "resolved_path": source, "size_bytes": 100, "sha256": "a" * 64, "binding": "reader_snapshot"}
    plan = {"mode": "manual", "total_pages": 4, "ranges": [[0, 2], [2, 4]], "entries": entries,
        "coverage": {"complete": True, "covered_pages": 4, "section_count": 2}, "normalized_inputs": {"starts": [1, 3]},
        "source_identity": source_id, "warnings": [], "notices": []}
    rid = "d" * 32
    final = base + "/processing_20000101_000000_" + rid
    outputs = [{"filename": entry["filename"], "range": [entry["start"], entry["end"]],
        "page_ids": list(range(entry["start"] + 1, entry["end"] + 1)), "sha256": "e" * 64} for entry in entries]
    written = [{**entry, "sha256": "e" * 64, "size_bytes": 100, "page_count": 2} for entry in entries]
    execution = {"run_id": rid, "final_directory": final, "mode": "manual", "total_pages": 4,
        "source_identity": source_id, "coverage": deepcopy(plan["coverage"]), "written_count": 2, "outputs": written,
        "manifest_filename": "WinBookSplit_Manifest.json"}
    execution["manifest"] = {"schema_version": 1, "status": "complete", **{key: deepcopy(execution[key]) for key in
        ("run_id", "final_directory", "mode", "total_pages", "source_identity", "coverage", "written_count", "outputs")}}
    frame = {"protocol": "winbooksplit.result", "version": 1, "status": "success", "code": "split_complete", "mode": "manual",
        "exit_code": 0, "written_count": 2, "warnings": [], "fallback_modes": [], "plan": plan, "execution": execution}
    outcome = {"protocol": "winbooksplit.outcome", "version": 1, "status": "success", "code": "split_complete", "mode": "manual",
        "exit_code": 0, "written_count": 2, "final_directory": final, "engine_result": frame}
    invocation = {"path": "C:/Python/python.exe", "arguments": ["-I", "-B", "-X", "utf8", app + "/engine/winbooksplit_engine.py",
        source, base, "manual", "1,3"]}
    summary = {"JobAssigned": True, "ParentStopped": True, "DescendantsStopped": True, "StreamsComplete": True,
        "TimedOut": False, "Cancelled": False, "Pid": 1234, "ExitCode": 0, "ResultRecordCount": 1,
        "StartError": None, "StreamError": None, "StopError": None, "ResultError": None}
    after = file_record(source, hashed="b") if kind == "replace" else {"path": source, "exists": False}
    fault = {"protocol": "winbooksplit.authored-source-fault", "version": 1, "state": "engine_completed", "kind": kind,
        "pid": 1234, "parent_pid": 5678, "python": "C:/Python/python.exe", "argv": invocation["arguments"][5:], "saved_engine_sha256": "a" * 64,
        "source_reader_calls": 1, "source_reread_attempts": [], "before_fault": deepcopy(before), "after_fault": deepcopy(after),
        "engine_main_return": None, "engine_system_exit_code": 0, "prepared_plan": deepcopy(plan),
        "staged_output_reader_paths": [base + "/.WinBookSplit-stage-" + rid + "/" + entry["filename"] for entry in entries]}
    prior_dir = base + "/immutable_20000101_000000_" + "f" * 32
    prior = {"path": prior_dir, "device": 1, "inode": 99, "members": {name: file_record(prior_dir + "/" + name, inode=n)
        for n, name in enumerate((".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", *(entry["filename"] for entry in entries)), 1)}}
    protected = {"original": file_record(books + "/immutable-original.pdf"), "replacement": file_record(books + "/replacement.pdf", hashed="b"),
        "neighbor": file_record(books + "/neighbor.txt", hashed="c"), "prior": prior}
    application = {name: "a" * 64 for name in validator.APPLICATION}
    adapter = validator.adapter_sha256()
    copied = {**application, "engine/winbooksplit_engine.py": adapter}
    parameters = ["-InputFile", source, "-OutputDirectory", base, "-PythonPath", "C:/Python/python.exe", "-Mode", "Manual",
        "-StartPages", "1,3", "-NonInteractive", "-NoPause"]
    case = {"id": host + "-captured-" + kind, "host_id": host, "kind": kind, "passed": True,
        "shell_executable": observation["shell_executable"], "host_version": observation["host_version"],
        "command": [observation["shell_executable"], "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-File",
            app + "/WinBookSplit.ps1", *parameters], "parameters": parameters, "cwd": cwd, "output_base": base,
        "python_executable": "C:/Python/python.exe", "actual_process": True, "pid": 5678, "exit_code": 0, "timed_out": False,
        "stdin": "DEVNULL", "stdin_utf8": None, "elapsed_seconds": 1.2, "timeout_seconds": 60,
        "processing_copy_before": before, "processing_copy_after": after, "processing_copy_after_cleanup": deepcopy(after),
        "source_observations_before": protected, "source_observations_after": deepcopy(protected),
        "source_observations_after_cleanup": {key: deepcopy(protected[key]) for key in ("original", "replacement", "neighbor")},
        "source_readonly_observed": True, "application_sha256": application, "copied_application_sha256": copied,
        "saved_engine_sha256": "a" * 64, "adapter_sha256": adapter, "source_fault": fault, "outcome": outcome,
        "engine_records": [frame], "engine_invocations": [invocation], "process_summaries": [summary], "outputs": outputs,
        "source_page_ids": [1, 2, 3, 4], "source_content_sha256": ["a" * 64, "b" * 64, "c" * 64, "d" * 64],
        "output_content_sha256": ["a" * 64, "b" * 64, "c" * 64, "d" * 64], "owned_outputs_removed": True,
        "marked_stage_absent": True, "prior_removed_after_preservation_proof": True,
        "prior_setup": {"actual_engine_api": True, "native_launcher_case": False, "execution": {
            "source_identity": {"sha256": "a" * 64}, "final_directory": prior_dir}, "outputs": deepcopy(outputs)}}
    publication = {"path": final, "device": 1, "inode": 999, "members": {name: file_record(final + "/" + name, hashed="e", inode=n)
        for n, name in enumerate((".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", *(entry["filename"] for entry in entries)), 1)}}
    owner_raw = json.dumps({"schema_version": 1, "kind": "run", "run_id": rid})
    manifest_raw = json.dumps(execution["manifest"])
    for name, raw in ((".WinBookSplit-owner.json", owner_raw), ("WinBookSplit_Manifest.json", manifest_raw)):
        publication["members"][name].update(size_bytes=len(raw.encode()), sha256=sha256(raw.encode()).hexdigest())
    case.update(publication_observation=publication, publication_owner_raw=owner_raw, publication_manifest_raw=manifest_raw)
    reseal(case)
    return case


def reseal(case):
    """Hash authored changed records independently of their semantic validity."""
    raw = json.dumps(case["source_fault"], separators=(",", ":")) + "\n"
    case.update(source_fault_raw=raw, source_fault_sha256=sha256(raw.encode()).hexdigest())
    case["source_fault_receipt"] = {"path": case["cwd"] + "/../source-fault.json", "available": True,
        "raw_text": raw, "sha256": case["source_fault_sha256"], "size_bytes": len(raw.encode())}
    text = "[ENGINE] " + json.dumps(case["engine_invocations"][0]) + "\r\n[PROCESS] " + json.dumps(case["process_summaries"][0])
    text += "\r\n" + json.dumps(case["engine_records"][0]) + "\r\n[OPERATION-OUTCOME] " + json.dumps(case["outcome"]) + "\r\n"
    case.update(console_log=text, console_log_sha256=sha256(text.encode()).hexdigest())
    case["stdout"] = "[OUTCOME] " + json.dumps(case["outcome"]) + "\n"
    case["stderr"] = ""
    for field in ("stdout", "stderr"):
        case[field + "_sha256"] = sha256(case[field].encode()).hexdigest()
    case["console_evidence"] = fixture.make(case["outcome"], case["console_log_sha256"], log_text=text,
        source_path=case["processing_copy_before"]["path"], source_observation={"sha256": "a" * 64, "size_bytes": 100})
    case["console_evidence"]["run_manifest"]["settings"]["non_interactive"] = True
    fixture.reseal(case["console_evidence"])


def source_map():
    return {name: "a" * 64 for name in validator.APPLICATION | {"tests/faults/cross_host.py", "tests/faults/validate_cross_host_report.py"}}


def complete_report():
    source = source_map()
    return {"protocol": "winbooksplit.fault-cross-host", "version": 1, "task_id": "M4-T03", "acceptance_ids": ["AC-075"],
        "result": "CROSS_HOST_SNAPSHOT_PASSED", "success": True, "exit_code": 0, "case_count": 4,
        "tested_path_sha256": source, "source_after_sha256": deepcopy(source), "source_unchanged": True,
        "host_cases": [host_record(host) for host in validator.HOSTS],
        "cases": [complete_case(host, kind) for host in validator.HOSTS for kind in validator.KINDS],
        "machine_settings_unchanged": True, "owned_outputs_removed": True, "fixture_workspace_owned_by_parent": True}


class CrossHostReceiptTests(unittest.TestCase):
    def test_complete_synthetic_four_case_contract_is_accepted(self):
        report = complete_report()
        validator.validate_cross_host_report(report, [host["shell_executable"] for host in report["host_cases"]])

    def check_rejected(self, mutate, *, reseal_records=False):
        case = complete_case()
        mutate(case)
        if reseal_records:
            reseal(case)
        with self.assertRaises(ValueError):
            validator.validate_case(case, host_record(), source_map())

    def test_resealed_captured_plan_and_source_hash_contradictions_reject(self):
        for mutate in (lambda c: c["source_fault"]["prepared_plan"].update(total_pages=5),
                lambda c: c["source_fault"]["prepared_plan"]["source_identity"].update(sha256="b" * 64),
                lambda c: c["source_fault"].update(source_reader_calls=2),
                lambda c: c["source_fault"].update(source_reread_attempts=[c["processing_copy_before"]["path"]]),
                lambda c: c["source_fault"].update(source_reader_calls=True),
                lambda c: c["source_fault"].update(staged_output_reader_paths=[])):
            with self.subTest(mutate=mutate):
                self.check_rejected(mutate, reseal_records=True)

    def test_resealed_processing_fault_must_actually_change_or_delete_copy(self):
        for mutate in (lambda c: c["processing_copy_after"].update(sha256="a" * 64),
                lambda c: c["source_fault"]["after_fault"].update(exists=False),
                lambda c: c["processing_copy_before"].update(size_bytes=1),
                lambda c: c["source_fault"].update(state="prepared_source_fault_applied")):
            with self.subTest(mutate=mutate):
                self.check_rejected(mutate, reseal_records=True)
        case = complete_case(kind="delete")
        case["processing_copy_after"] = file_record(case["processing_copy_before"]["path"])
        case["source_fault"]["after_fault"] = deepcopy(case["processing_copy_after"])
        reseal(case)
        with self.assertRaises(ValueError):
            validator.validate_case(case, host_record(), source_map())

    def test_native_command_streams_exit_and_no_prompt_are_required(self):
        for mutate in (lambda c: c.update(exit_code=7), lambda c: c.update(stdin="PIPE"),
                lambda c: c["command"].__setitem__(0, "C:/other/powershell.exe"),
                lambda c: c["command"].__setitem__(-2, "-Preview"),
                lambda c: c.update(stdout=c["stdout"] + "Create these chapter PDFs?"),
                lambda c: c.update(stdout_sha256="e" * 64), lambda c: c.update(elapsed_seconds=float("inf"))):
            with self.subTest(mutate=mutate):
                self.check_rejected(mutate)

    def test_resealed_shutdown_missing_record_and_child_pid_contradictions_reject(self):
        for mutate in (lambda c: c["process_summaries"][0].update(StreamsComplete=False),
                lambda c: c["process_summaries"][0].update(DescendantsStopped=False),
                lambda c: c["process_summaries"][0].update(ResultRecordCount=0),
                lambda c: c["process_summaries"][0].update(Pid=9999),
                lambda c: c["source_fault"].update(engine_system_exit_code=7)):
            with self.subTest(mutate=mutate):
                self.check_rejected(mutate, reseal_records=True)

    def test_actual_adapter_maps_none_return_without_inventing_native_status(self):
        harness = load("wbs_snapshot_adapter_return_unit", ROOT / "tests/faults/cross_host.py")
        body = ast.parse(harness.ADAPTER).body
        start = next(index for index, node in enumerate(body) if isinstance(node, ast.Assign)
            and isinstance(node.targets[0], ast.Name) and node.targets[0].id == "result")
        portion = compile(ast.Module(body=body[start:], type_ignores=[]), "actual-owned-adapter-return", "exec")
        for main_return, intended in ((None, 0), (0, 0), (19, 19)):
            recorded = {}
            namespace = {"engine": SimpleNamespace(main=lambda: main_return), "observation": recorded, "save": lambda: None}
            with self.subTest(main_return=main_return), self.assertRaises(SystemExit) as raised:
                exec(portion, namespace)
            self.assertIs(raised.exception.code, main_return)
            self.assertEqual(recorded["engine_main_return"], main_return)
            self.assertEqual(recorded["engine_system_exit_code"], intended)
            self.assertNotIn("engine_native_exit", recorded)

    def test_resealed_direct_venv_redirector_parent_link_is_explicit_and_required(self):
        case = complete_case()
        case["process_summaries"][0]["Pid"] = 6789
        case["source_fault"]["parent_pid"] = 6789
        reseal(case)
        validator.validate_case(case, host_record(), source_map())
        for value in (None, True, 0, 4321):
            changed = deepcopy(case)
            changed["source_fault"]["parent_pid"] = value
            reseal(changed)
            with self.subTest(parent_pid=value), self.assertRaises(ValueError):
                validator.validate_case(changed, host_record(), source_map())
        changed = deepcopy(case)
        del changed["source_fault"]["parent_pid"]
        reseal(changed)
        with self.assertRaises(ValueError):
            validator.validate_case(changed, host_record(), source_map())

    def test_protected_original_neighbor_prior_and_owned_cleanup_are_bound(self):
        for mutate in (lambda c: c["source_observations_after"]["neighbor"].update(sha256="e" * 64),
                lambda c: c["source_observations_after"]["prior"]["members"]["WinBookSplit_Manifest.json"].update(inode=888),
                lambda c: c["source_observations_after_cleanup"]["original"].update(inode=888),
                lambda c: c.update(owned_outputs_removed=False), lambda c: c.update(marked_stage_absent=False),
                lambda c: c.update(prior_removed_after_preservation_proof=False)):
            with self.subTest(mutate=mutate):
                self.check_rejected(mutate)

    def test_saved_engine_and_declared_adapter_bytes_are_required(self):
        for mutate in (lambda c: c.update(saved_engine_sha256="c" * 64),
                lambda c: c["copied_application_sha256"].update({"WinBookSplit.ps1": "c" * 64}),
                lambda c: c.update(adapter_sha256="c" * 64)):
            with self.subTest(mutate=mutate):
                self.check_rejected(mutate)

    def test_resealed_publication_member_owner_and_manifest_bytes_are_bound(self):
        for mutate in (lambda c: c["publication_observation"]["members"]["01 - Section (Page 1-2).pdf"].update(sha256="c" * 64),
                lambda c: c["publication_observation"]["members"]["02 - Section (Page 3-4).pdf"].update(size_bytes=1),
                lambda c: c["publication_observation"].update(path="C:/other/run"),
                lambda c: c["publication_observation"]["members"].update({"foreign.txt": file_record("C:/other/foreign.txt")}),
                lambda c: c.update(publication_owner_raw=json.dumps({"schema_version": 1, "kind": "run", "run_id": "e" * 32})),
                lambda c: c.update(publication_manifest_raw=json.dumps({"schema_version": 1, "status": "complete"}))):
            with self.subTest(mutate=mutate):
                self.check_rejected(mutate, reseal_records=True)

    def test_resealed_reopened_page_identity_and_content_loss_reject(self):
        for mutate in (lambda c: c["outputs"][0].update(page_ids=[2, 1]),
                lambda c: c["outputs"][1].update(page_ids=[2, 3]),
                lambda c: c.update(output_content_sha256=list(reversed(c["source_content_sha256"]))),
                lambda c: c["outcome"].update(written_count=0), lambda c: c["outcome"].update(version=True)):
            with self.subTest(mutate=mutate):
                self.check_rejected(mutate, reseal_records=True)

    def test_exact_case_hosts_source_map_and_typed_aggregate_required(self):
        valid = complete_report()
        shells = [host["shell_executable"] for host in valid["host_cases"]]
        for mutate in (lambda r: r["cases"].pop(), lambda r: r["cases"].__setitem__(1, deepcopy(r["cases"][0])),
                lambda r: r["host_cases"].pop(), lambda r: r.update(case_count=True), lambda r: r.update(exit_code=False),
                lambda r: r["source_after_sha256"].update({"WinBookSplit.ps1": "c" * 64})):
            report = deepcopy(valid)
            mutate(report)
            with self.subTest(mutate=mutate), self.assertRaises(ValueError):
                validator.validate_cross_host_report(report, shells)

    def test_cleanup_rejects_unexpected_or_replaced_authenticated_member_before_deletion(self):
        harness = load("wbs_cross_host_cleanup_unit", ROOT / "tests/faults/cross_host.py")
        with tempfile.TemporaryDirectory(prefix="wbs-snapshot-cleanup-") as work:
            base = Path(work).resolve()
            directory = base / ("source_20000101_000000_" + "a" * 32)
            directory.mkdir()
            for name in (".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", "authored.pdf"):
                (directory / name).write_bytes(b"Authored non-PDF cleanup unit, no child\n")
            execution = {"run_id": "a" * 32, "final_directory": str(directory), "outputs": [{"filename": "authored.pdf"}]}
            expected = harness.tree_identity(directory)
            unexpected = directory / "foreign-neighbor.txt"
            unexpected.write_bytes(b"Must not authorize deleting this unregistered member\n")
            with patch.object(harness.launchers, "remove_known_directory") as remove:
                with self.assertRaisesRegex((ValueError, RuntimeError), "unexpected or changed"):
                    harness.remove_run(base, execution, expected)
                remove.assert_not_called()
            unexpected.unlink()
            (directory / "authored.pdf").write_bytes(b"Changed bytes after authentication\n")
            with patch.object(harness.launchers, "remove_known_directory") as remove:
                with self.assertRaisesRegex((ValueError, RuntimeError), "unexpected or changed"):
                    harness.remove_run(base, execution, expected)
                remove.assert_not_called()


if __name__ == "__main__":
    unittest.main(verbosity=2)
