"""Strict AC-075 captured-source receipt checks; no historical-cause inference."""

from hashlib import sha256
import ast
import importlib.util
import json
import math
import ntpath
from pathlib import Path, PureWindowsPath
import re

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("wbs_snapshot_console_guard", ROOT / "tests/manual/current_launchers.py")
console = importlib.util.module_from_spec(spec)
spec.loader.exec_module(console)
spec = importlib.util.spec_from_file_location("wbs_snapshot_cli_guard", ROOT / "tests/cli/validate_cli_report.py")
cli = importlib.util.module_from_spec(spec)
spec.loader.exec_module(cli)
HOSTS = ("PS51", "PS7")
KINDS = ("replace", "delete")
APPLICATION = cli.APPLICATION


def adapter_sha256():
    parsed = ast.parse((ROOT / "tests/faults/cross_host.py").read_text(encoding="utf-8-sig"))
    values = [ast.literal_eval(node.value) for node in parsed.body if isinstance(node, ast.Assign)
        and any(isinstance(target, ast.Name) and target.id == "ADAPTER" for target in node.targets)]
    need(len(values) == 1 and type(values[0]) is str, "One literal declared copied-engine adapter required")
    return sha256(values[0].encode("utf-8")).hexdigest()


def need(value, message):
    if not value:
        raise ValueError(message)


def absolute(value):
    need(type(value) is str and PureWindowsPath(value).is_absolute(), "Absolute authored Windows path missing")
    return ntpath.normcase(ntpath.normpath(value))


def hashes(value):
    return type(value) is dict and bool(value) and all(type(key) is str and type(item) is str
        and re.fullmatch(r"[a-f0-9]{64}", item) for key, item in value.items())


def identity(record, *, exists=True):
    need(type(record) is dict and record.get("exists") is exists, "Authored file existence observation differs")
    absolute(record.get("path"))
    if not exists:
        need(set(record) == {"path", "exists"}, "Deleted processing path invented file contents/identity")
        return
    need(set(record) == {"path", "exists", "device", "inode", "attributes", "size_bytes", "sha256"}
        and all(type(record.get(key)) is int and record[key] >= 0 for key in ("device", "inode", "attributes", "size_bytes"))
        and record["inode"] > 0 and record["size_bytes"] > 0 and hashes({"file": record.get("sha256")}),
        "Concrete ordinary file hash/identity missing")


def tree(record):
    need(type(record) is dict and set(record) == {"path", "device", "inode", "members"}
        and type(record.get("device")) is int and record["device"] >= 0
        and type(record.get("inode")) is int and record["inode"] > 0
        and type(record.get("members")) is dict, "Concrete prior publication identity missing")
    parent = absolute(record.get("path"))
    need(set(record["members"]) == {".WinBookSplit-owner.json", "WinBookSplit_Manifest.json",
        "01 - Section (Page 1-2).pdf", "02 - Section (Page 3-4).pdf"}, "Prior publication exact membership missing")
    for name, member in record["members"].items():
        identity(member)
        need(absolute(member["path"]) == ntpath.join(parent, name.lower()), "Prior member path escaped its publication")


def source_observations(value):
    need(type(value) is dict and set(value) == {"original", "replacement", "neighbor", "prior"},
        "Original/replacement/neighbor/prior observations missing")
    for name in ("original", "replacement", "neighbor"):
        identity(value[name])
    tree(value["prior"])


def fault_binding(case):
    fault, raw = case.get("source_fault"), case.get("source_fault_raw")
    need(type(raw) is str and type(fault) is dict and json.loads(raw) == fault
        and sha256(raw.encode("utf-8")).hexdigest() == case.get("source_fault_sha256"), "Actual source fault raw/hash differs")
    captured = case.get("source_fault_receipt")
    need(type(captured) is dict and captured.get("available") is True and captured.get("raw_text") == raw
        and captured.get("sha256") == case["source_fault_sha256"] and captured.get("size_bytes") == len(raw.encode("utf-8")),
        "Prepared source fault raw file capture differs")
    need(fault.get("protocol") == "winbooksplit.authored-source-fault" and type(fault.get("version")) is int
        and fault["version"] == 1 and fault.get("state") == "engine_completed"
        and fault.get("kind") == case.get("kind") and type(fault.get("pid")) is int and fault["pid"] > 0
        and type(fault.get("parent_pid")) is int and fault["parent_pid"] > 0
        and "engine_main_return" in fault and (fault["engine_main_return"] is None or
            type(fault["engine_main_return"]) is int and fault["engine_main_return"] == 0)
        and type(fault.get("engine_system_exit_code")) is int and fault["engine_system_exit_code"] == 0,
        "Source fault did not complete through the real engine")
    need(type(fault.get("source_reader_calls")) is int and fault["source_reader_calls"] == 1
        and fault.get("source_reread_attempts") == [], "Engine reopened/replanned its stale source")
    need(fault.get("before_fault") == case.get("processing_copy_before")
        and fault.get("after_fault") == case.get("processing_copy_after"), "Prepared source fault identity differs from actual file observations")
    identity(case.get("processing_copy_before"))
    identity(case.get("processing_copy_after"), exists=case["kind"] != "delete")
    before = case["processing_copy_before"]
    immutable = case["source_observations_before"]
    need(before["sha256"] == immutable["original"]["sha256"] and before["size_bytes"] == immutable["original"]["size_bytes"],
        "Processing copy did not start with the immutable authored PDF")
    need(immutable["original"]["sha256"] != immutable["replacement"]["sha256"], "Reverse-page replacement did not differ")
    if case["kind"] == "replace":
        need(case["processing_copy_after"]["sha256"] == immutable["replacement"]["sha256"]
            and case["processing_copy_after"]["size_bytes"] == immutable["replacement"]["size_bytes"], "Replacement fault was not actually applied")
    need(fault.get("saved_engine_sha256") == case.get("saved_engine_sha256")
        and absolute(fault.get("python")) == absolute(case.get("python_executable")), "Real saved engine/runtime byte binding differs")


def native_binding(case, host):
    parameters = ["-InputFile", case["processing_copy_before"]["path"], "-OutputDirectory", case.get("output_base"),
        "-PythonPath", case.get("python_executable"), "-Mode", "Manual", "-StartPages", "1,3", "-NonInteractive", "-NoPause"]
    command = case.get("command")
    need(type(command) is list and len(command) == 19 and all(type(value) is str for value in command)
        and command[0] == host["shell_executable"] and command[1:6] ==
        ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-File"]
        and ntpath.basename(absolute(command[6])) == "winbooksplit.ps1"
        and command[7:] == parameters == case.get("parameters"), "Actual host/application literal argv differs")
    app = ntpath.dirname(absolute(command[6]))
    cwd = absolute(case.get("cwd"))
    base = absolute(case.get("output_base"))
    need(ntpath.basename(cwd) == "c" and ntpath.basename(base) == "o" and ntpath.basename(app) == "a"
        and ntpath.dirname(cwd) == ntpath.dirname(base) == ntpath.dirname(app), "Authored application/cwd/output scope differs")
    need(case.get("actual_process") is True and type(case.get("pid")) is int and case["pid"] > 0
        and type(case.get("exit_code")) is int and case["exit_code"] == 0 and case.get("timed_out") is False
        and case.get("stdin") == "DEVNULL" and case.get("stdin_utf8") is None
        and type(case.get("elapsed_seconds")) in (int, float) and math.isfinite(case["elapsed_seconds"])
        and 0 <= case["elapsed_seconds"] < case.get("timeout_seconds", 0) == 60, "Actual native status/stdin/deadline missing")
    for field in ("stdout", "stderr"):
        text = case.get(field)
        need(type(text) is str and case.get(field + "_sha256") == sha256(text.encode("utf-8")).hexdigest(),
            "Actual native stream/hash differs")
    need(not any(text in case["stdout"] + case["stderr"] for text in
        ("Create these chapter PDFs?", "Open output folder?", "Open the generated PDF?", "Read-Host", "[FALLBACK]")),
        "Noninteractive captured-source case prompted/opened/fell back")
    final = case.get("outcome")
    lines = [json.loads(line[len("[OUTCOME] "):]) for line in case["stdout"].splitlines() if line.startswith("[OUTCOME] ")]
    need(type(final) is dict and lines == [final] and final.get("protocol") == "winbooksplit.outcome"
        and type(final.get("version")) is int and final["version"] == 1
        and (final.get("status"), final.get("code"), final.get("mode"), final.get("exit_code"), final.get("written_count"))
            == ("success", "split_complete", "manual", 0, 2)
        and all(type(final.get(key)) is int for key in ("exit_code", "written_count")), "Actual sole terminal outcome contradicts the operation")
    frame = final.get("engine_result")
    need(type(frame) is dict and frame.get("protocol") == "winbooksplit.result" and type(frame.get("version")) is int
        and frame["version"] == 1 and all(frame.get(key) == final.get(key) for key in ("status", "code", "mode", "exit_code", "written_count"))
        and type(frame.get("exit_code")) is int and type(frame.get("written_count")) is int
        and frame.get("warnings") == [] and frame.get("fallback_modes") == [] and case.get("engine_records") == [frame],
        "Actual engine result/terminal status differs")
    plan, execution = frame.get("plan"), frame.get("execution")
    need(type(plan) is dict and plan == case["source_fault"].get("prepared_plan"), "Execution differs from the single captured prepared plan")
    cli.coverage(plan, [[0, 2], [2, 4]])
    need(type(plan["coverage"].get("complete")) is bool and plan["coverage"]["complete"] is True
        and all(type(plan["coverage"].get(key)) is int for key in ("covered_pages", "section_count"))
        and all(type(entry.get("sequence")) is int for entry in plan["entries"]), "Plan typed coverage/sequence missing")
    need(plan["total_pages"] == 4 and plan.get("mode") == "manual" and plan.get("normalized_inputs") == {"starts": [1, 3]},
        "Captured manual normalized plan differs")
    identity_source = plan.get("source_identity", {})
    before = case["processing_copy_before"]
    need(identity_source.get("path") == before["path"] and identity_source.get("sha256") == before["sha256"]
        and identity_source.get("size_bytes") == before["size_bytes"] and identity_source.get("binding") == "reader_snapshot",
        "Captured plan source SHA/identity differs from before the fault")
    need(type(execution) is dict and execution.get("source_identity") == identity_source and execution.get("coverage") == plan["coverage"]
        and execution.get("total_pages") == 4 and execution.get("mode") == "manual"
        and type(execution.get("written_count")) is int and execution["written_count"] == 2
        and execution.get("final_directory") == final.get("final_directory"), "Publication source/plan/count differs")
    run_id = execution.get("run_id")
    need(type(run_id) is str and re.fullmatch(r"[0-9a-f]{32}", run_id)
        and ntpath.dirname(absolute(execution.get("final_directory"))) == base, "Publication escaped its direct base")
    outputs = case.get("outputs")
    need(type(outputs) is list and len(outputs) == 2 and [item.get("range") for item in outputs] == [[0, 2], [2, 4]]
        and [item.get("page_ids") for item in outputs] == [[1, 2], [3, 4]]
        and case.get("source_page_ids") == [1, 2, 3, 4]
        and case.get("source_content_sha256") == case.get("output_content_sha256")
        and type(case.get("source_content_sha256")) is list and len(case["source_content_sha256"]) == 4,
        "Reopened outputs lost/reordered original physical pages/content")
    need([item.get("filename") for item in outputs] == [entry["filename"] for entry in plan["entries"]]
        == [item.get("filename") for item in execution.get("outputs", [])], "Output names differ from the prepared plan")
    for output, entry in zip(outputs, execution["outputs"]):
        need(hashes({"pdf": output.get("sha256")}) and output["sha256"] == entry.get("sha256")
            and type(entry.get("size_bytes")) is int and entry["size_bytes"] > 0 and entry.get("page_count") == 2,
            "Reopened output bytes/count differ from completion manifest")
    expected_manifest = {"schema_version": 1, "status": "complete", **{key: execution[key] for key in
        ("run_id", "final_directory", "mode", "total_pages", "source_identity", "coverage", "written_count", "outputs")}}
    need(execution.get("manifest") == expected_manifest and execution.get("manifest_filename") == "WinBookSplit_Manifest.json",
        "Actual output completion manifest differs from execution")
    publication = case.get("publication_observation")
    tree(publication)
    need(publication["path"] == execution["final_directory"], "Authenticated publication directory differs from terminal execution")
    for raw_key, name, expected in (("publication_owner_raw", ".WinBookSplit-owner.json",
            {"schema_version": 1, "kind": "run", "run_id": run_id}),
            ("publication_manifest_raw", "WinBookSplit_Manifest.json", expected_manifest)):
        raw = case.get(raw_key)
        member = publication["members"][name]
        need(type(raw) is str and json.loads(raw) == expected
            and sha256(raw.encode("utf-8")).hexdigest() == member["sha256"]
            and len(raw.encode("utf-8")) == member["size_bytes"], "Authenticated publication owner/manifest raw bytes differ")
    for entry in execution["outputs"]:
        member = publication["members"][entry["filename"]]
        need(member["sha256"] == entry["sha256"] and member["size_bytes"] == entry["size_bytes"],
            "Authenticated publication member bytes differ from completion manifest/reopened PDF")
    stages = case["source_fault"].get("staged_output_reader_paths")
    need(type(stages) is list and len(stages) == 2 and {absolute(value) for value in stages}
        == {ntpath.join(base, ".winbooksplit-stage-" + run_id, entry["filename"].lower()) for entry in plan["entries"]},
        "Real engine did not validate both staged outputs without reopening the source")
    summaries = case.get("process_summaries")
    need(type(summaries) is list and len(summaries) == 1, "Actual supervisor summary missing/duplicated")
    summary = summaries[0]
    need(all(summary.get(key) is True for key in ("ParentStopped", "DescendantsStopped", "StreamsComplete", "JobAssigned"))
        and summary.get("TimedOut") is False and summary.get("Cancelled") is False
        and type(summary.get("Pid")) is int and summary["Pid"] > 0
        and (summary["Pid"] == case["source_fault"]["pid"] or summary["Pid"] == case["source_fault"]["parent_pid"])
        and type(summary.get("ExitCode")) is int and summary["ExitCode"] == 0
        and type(summary.get("ResultRecordCount")) is int and summary["ResultRecordCount"] == 1
        and all(key in summary and summary[key] is None for key in ("StartError", "StreamError", "StopError", "ResultError")),
        "Actual native child/tree/streams/result shutdown is unproved")
    invocations = case.get("engine_invocations")
    expected_engine = ["-I", "-B", "-X", "utf8", ntpath.join(app, "engine", "winbooksplit_engine.py"),
        absolute(before["path"]), base, "manual", "1,3"]
    need(type(invocations) is list and len(invocations) == 1 and absolute(invocations[0].get("path")) == absolute(case["python_executable"])
        and [absolute(value) if index in {4, 5, 6} else value for index, value in enumerate(invocations[0].get("arguments", []))] == expected_engine
        and case["source_fault"].get("argv") == invocations[0]["arguments"][5:], "Actual engine argv/interpreter differs")
    log = case.get("console_log")
    need(type(log) is str and sha256(log.encode("utf-8")).hexdigest() == case.get("console_log_sha256")
        and [json.loads(line[len("[PROCESS] "):]) for line in log.splitlines() if line.startswith("[PROCESS] ")] == summaries
        and [json.loads(line[len("[ENGINE] "):]) for line in log.splitlines() if line.startswith("[ENGINE] ")] == invocations
        and [json.loads(line) for line in log.splitlines() if line.startswith('{')] == [frame], "Raw console/engine/supervisor binding differs")
    console.validate_console_evidence(case.get("console_evidence"), final, case["console_log_sha256"],
        source_path=before["path"], source_observation={"sha256": before["sha256"], "size_bytes": before["size_bytes"]}, log_text=log)
    settings = case["console_evidence"]["run_manifest"]["settings"]
    need(settings.get("non_interactive") is True and settings.get("preview") is False
        and settings.get("no_pause") is True and settings.get("keep_converted_pdf") is False,
        "Local run settings contradict the actual complete noninteractive command")


def validate_case(case, host, source_sha256, *, cleanup_complete=True):
    need(type(case) is dict and case.get("passed") is True and case.get("host_id") == host["id"]
        and case.get("kind") in KINDS and case.get("id") == host["id"] + "-captured-" + case["kind"], "Wrong captured-source case")
    need(case.get("shell_executable") == host["shell_executable"] and case.get("host_version") == host["host_version"],
        "Captured-source case used another host")
    source_observations(case.get("source_observations_before"))
    source_observations(case.get("source_observations_after"))
    need(case["source_observations_before"] == case["source_observations_after"] and case.get("source_readonly_observed") is True,
        "Original/replacement/neighbor/prior changed during the authored processing-copy fault")
    application, copied = case.get("application_sha256"), case.get("copied_application_sha256")
    need(hashes(application) and hashes(copied) and set(application) == set(copied) == APPLICATION
        and all(application[name] == source_sha256.get(name) for name in application)
        and all(copied[name] == application[name] for name in APPLICATION - {"engine/winbooksplit_engine.py"})
        and case.get("saved_engine_sha256") == application["engine/winbooksplit_engine.py"]
        and case.get("adapter_sha256") == copied["engine/winbooksplit_engine.py"] == adapter_sha256(), "Copied application/seam/saved engine raw byte binding differs")
    fault_binding(case)
    native_binding(case, host)
    prior = case.get("prior_setup")
    need(type(prior) is dict and prior.get("actual_engine_api") is True and prior.get("native_launcher_case") is False
        and type(prior.get("execution")) is dict and prior["execution"].get("source_identity", {}).get("sha256")
            == case["source_observations_before"]["original"]["sha256"]
        and prior["execution"].get("final_directory") == case["source_observations_before"]["prior"]["path"]
        and [item.get("page_ids") for item in prior.get("outputs", [])] == [[1, 2], [3, 4]], "Genuine prior publication precondition missing")
    if cleanup_complete:
        need(case.get("owned_outputs_removed") is True and case.get("marked_stage_absent") is True
            and case.get("prior_removed_after_preservation_proof") is True, "Exact marked output cleanup is unproved")
        need(case.get("source_observations_after_cleanup") == {key: case["source_observations_before"][key]
            for key in ("original", "replacement", "neighbor")}
            and case.get("processing_copy_after_cleanup") == case["processing_copy_after"], "Own cleanup changed protected inputs")


def validate_cross_host_report(report, requested_shell_paths, *, cleanup_complete=True, expected_source_sha256=None):
    need(type(report) is dict and report.get("protocol") == "winbooksplit.fault-cross-host"
        and type(report.get("version")) is int and report["version"] == 1 and report.get("task_id") == "M4-T03"
        and report.get("acceptance_ids") == ["AC-075"] and report.get("result") == "CROSS_HOST_SNAPSHOT_PASSED"
        and report.get("success") is True and type(report.get("exit_code")) is int and report["exit_code"] == 0,
        "Wrong cross-host receipt type/status")
    source = report.get("tested_path_sha256")
    need(hashes(source) and report.get("source_after_sha256") == source and report.get("source_unchanged") is True,
        "Actual cross-host sources changed/missing")
    need(APPLICATION | {"tests/faults/cross_host.py", "tests/faults/validate_cross_host_report.py"} <= set(source),
        "Cross-host application/adapter/guard source hashes missing")
    if expected_source_sha256 is not None:
        need(source == expected_source_sha256, "Cross-host source map differs from the parent fault stage")
    hosts = report.get("host_cases")
    need(type(hosts) is list and len(hosts) == 2 and {host.get("id") for host in hosts} == set(HOSTS)
        and len(requested_shell_paths) == 2 and {absolute(value) for value in requested_shell_paths}
            == {absolute(host.get("shell_executable")) for host in hosts}, "Both requested actual hosts required without skipping")
    by_host = {host["id"]: host for host in hosts}
    for host in hosts:
        need(host.get("passed") is True and host.get("host_major") == (5 if host["id"] == "PS51" else 7)
            and type(host.get("host_version")) is str and host["host_version"].startswith("5.1." if host["id"] == "PS51" else "7.")
            and host.get("exit_code") == 0 and type(host.get("stored_policies")) is list and host["stored_policies"]
            and host["stored_policies"] == host.get("policies_after"), "Actual host/version/stored settings missing or changed")
    cases = report.get("cases")
    expected = {host + "-captured-" + kind for host in HOSTS for kind in KINDS}
    need(type(cases) is list and len(cases) == 4 and {case.get("id") for case in cases} == expected
        and type(report.get("case_count")) is int and report["case_count"] == 4, "Exact four captured-source cases required")
    for case in cases:
        validate_case(case, by_host[case["host_id"]], source, cleanup_complete=True)
    need(report.get("machine_settings_unchanged") is True and report.get("owned_outputs_removed") is True
        and report.get("fixture_workspace_owned_by_parent") is True and "owned_temp_removed" not in report,
        "Cross-host settings/marked-output cleanup versus parent workspace scope invalid")
