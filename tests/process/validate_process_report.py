"""Independent structural and byte-level M2-T05 receipt checks (not authentication)."""

from __future__ import annotations

import json
import math
import ntpath
from pathlib import PureWindowsPath
import re
from hashlib import sha256
import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location("wbs_process_interaction_validation", Path(__file__).resolve().parents[2] / "tests/manual/interaction_receipts.py")
interactions = importlib.util.module_from_spec(spec)
spec.loader.exec_module(interactions)
spec = importlib.util.spec_from_file_location("wbs_process_console_validation", Path(__file__).resolve().parents[2] / "tests/manual/current_launchers.py")
console = importlib.util.module_from_spec(spec)
spec.loader.exec_module(console)

ACCEPTANCE = ["AC-046", "AC-047", "AC-048", "AC-049"]
HOSTS = ("PS51", "PS7")
KINDS = ("flood", "fast-tail", "unicode-boundary", "arguments", "large-frame", "oversized-frame",
         "duplicate-frames", "no-frame", "malformed-frame", "timeout-tree", "cancel-tree", "inherited-pipe", "detached-pipe")
UNICODE = "Janne ääkköset Größe 日本 中文 한국어"
ARGUMENTS = ["", "space value", "[%!] & (literal)", "O'Brien", 'a"b', 'end\\',
             'slash\\\\"quote', UNICODE, "$(Set-Content injected.txt); __import__('os').system('echo injected')"]


def require(value, message):
    if not value:
        raise ValueError(message)


def digest(value):
    require(type(value) is str and re.fullmatch(r"[a-f0-9]{64}", value) is not None, "Invalid SHA256")


def path(value):
    require(type(value) is str and PureWindowsPath(value).is_absolute(), "Absolute Windows path required")
    return ntpath.normcase(ntpath.normpath(value))


def flags(record, *fields):
    for field in fields:
        require(record.get(field) is True, field + ": absent or contradictory proof")


def exact_cases(records, expected):
    require(type(records) is list and all(type(item) is dict for item in records), "Missing cases")
    identifiers = [item.get("id") for item in records]
    require(len(identifiers) == len(expected) and set(identifiers) == set(expected), "Exact case set required, without duplicates")
    return {item["id"]: item for item in records}


def artifact(record, *, available=True):
    require(type(record) is dict and type(record.get("available")) is bool, "Authored raw artifact availability missing")
    path(record.get("path"))
    require(record["available"] is available, "Authored artifact availability contradicts observed readiness")
    if not available:
        require(not any(key in record for key in ("sha256", "raw_text", "size_bytes")), "Missing artifact invented raw evidence")
        return None
    for key in ("size_bytes", "created_ns", "modified_ns"):
        require(type(record.get(key)) is int and record[key] >= 0, "Authored raw artifact stat missing")
    raw = record.get("raw_text")
    require(type(raw) is str and record["size_bytes"] <= 16 * 1024 * 1024
            and len(raw.encode("utf-8")) == record["size_bytes"], "Authored raw artifact length differs")
    digest(record.get("sha256"))
    require(sha256(raw.encode("utf-8")).hexdigest() == record["sha256"], "Authored raw artifact hash differs")
    return raw


def native_host_capture(capture, host, *, complete=True):
    require(type(capture) is dict and capture.get("host_id") == host["id"], "Native host capture identity differs")
    command = capture.get("command")
    cwd = path(capture.get("cwd"))
    require(type(command) is list and len(command) == 11 and path(command[0]) == path(host["shell_executable"])
        and command[1:6] == ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-File"]
        and ntpath.basename(path(command[6])) == "probe-process.ps1"
        and command[7::2] == ["-PayloadPath", "-ReportPath"]
        and ntpath.dirname(path(command[8])) == cwd == ntpath.dirname(path(command[10])), "Native host exact argv/cwd missing")
    require(path(capture.get("payload", {}).get("path")) == path(command[8]), "Native payload path differs")
    payload = json.loads(artifact(capture.get("payload")))
    require(path(payload.get("work")) == cwd and type(payload.get("cases")) is list
        and [item.get("id") for item in payload["cases"]] == [host["id"] + "-" + kind for kind in KINDS],
        "Native authored payload/case set differs")
    require(ntpath.basename(path(payload.get("helper"))) == "winbooksplit.process.ps1"
        and ntpath.basename(path(payload.get("diagnostics"))) == "winbooksplit.diagnostics.ps1", "Native helper source scope differs")
    path(payload.get("python"))
    for item, expected_kind in zip(payload["cases"], KINDS):
        kind = item["kind"]
        arguments = item.get("arguments")
        require(kind == expected_kind and type(arguments) is list and len(arguments) >= 6
            and arguments[:4] == ["-I", "-B", "-X", "utf8"]
            and ntpath.basename(path(arguments[4])) == "native_child.py" and arguments[5] == kind,
            "Native authored child/kind argv differs")
        if kind == "arguments":
            require(arguments[6:] == ARGUMENTS, "Native literal argument recipe differs")
        elif kind in {"timeout-tree", "cancel-tree", "inherited-pipe", "detached-pipe"}:
            require([path(value) for value in arguments[6:]] == [ntpath.join(cwd, kind + "-parent.json"), ntpath.join(cwd, kind + "-child.json")],
                "Native authored PID readiness path scope differs")
        else:
            require(len(arguments) == 6, "Native child received undeclared extra arguments")
        require(type(item.get("timeout")) in (int, float)
                and item.get("timeout") == (1.0 if kind in {"timeout-tree", "inherited-pipe", "detached-pipe"} else 8.0)
                and item.get("cancel") is (kind == "cancel-tree"), "Native readiness/deadline recipe changed")
    process = capture.get("process")
    if not complete and process is None:
        error = capture.get("supervision_error")
        require(type(error) is dict and error.get("native_exit_available") is False
            and error.get("raw_streams_available") is False and error.get("inner_tree_shutdown_proved") is False,
            "Unproved native supervision invented exit/streams/shutdown")
        return None
    require(type(process) is dict and type(process.get("exit_code")) is int
        and type(process.get("stdout")) is str and type(process.get("stderr")) is str,
        "Actual native host exit/streams missing")
    raw_report = capture.get("native_report")
    require(path(raw_report.get("path")) == path(command[10]), "Native raw receipt path differs")
    if not complete and raw_report.get("available") is False:
        artifact(raw_report, available=False)
        return None
    native_report = json.loads(artifact(raw_report))
    if complete:
        require(process["exit_code"] == 0 and native_report.get("host_version") == host["host_version"]
            and native_report.get("stored_policies") == host["stored_policies"], "Native raw host failed or changed settings")
    return native_report


def validate_failed_process_report(report):
    """A retained failed attempt is evidence; this guard never turns it into PASS."""
    require(type(report) is dict and report.get("schema_version") == 1 and report.get("task_id") == "M2-T05"
        and report.get("result") == "PROCESS_REGRESSION_FAILED" and report.get("success") is False
        and type(report.get("exit_code")) is int and report["exit_code"] != 0,
        "Wrong failed process receipt type/status")
    require(report.get("cleanup_safe") is False and report.get("owned_temp_removed") is False,
        "Failed process evidence must remain retained without cleanup authority")
    path(report.get("workspace_retained"))
    require(type(report.get("error")) is str and (bool(report["error"]) or report.get("error_type") == "KeyboardInterrupt")
        and type(report.get("error_type")) is str,
        "Failed assertion/error missing")
    before, after = report.get("tested_path_sha256"), report.get("source_after_sha256")
    require(type(before) is dict and type(after) is dict and type(report.get("source_unchanged")) is bool
        and report["source_unchanged"] is (bool(before) and before == after), "Failed source observation contradicts maps")
    for value in [*before.values(), *after.values()]:
        digest(value)
    stage, case = report.get("failure_stage"), report.get("last_case")
    require(stage in {"setup", "native_host", "native_case", "application_case"}, "Failed stage missing")
    if stage == "setup":
        return
    require(type(case) is dict and case.get("passed") is False and type(case.get("command")) is list
        and bool(case["command"]), "Failed attempted case/argv missing")
    path(case.get("cwd"))
    if stage in {"native_host", "native_case"}:
        hosts = {item["id"]: item for item in report.get("host_cases", [])}
        captures = report.get("native_host_evidence")
        require(type(captures) is list and bool(captures), "Failed native raw host evidence missing")
        host_id = case.get("host_id") if stage == "native_host" else case.get("id", "").split("-")[0]
        require(host_id in hosts, "Failed actual host missing")
        matching = [item for item in captures if item.get("host_id") == host_id]
        require(len(matching) == 1, "Failed native host capture not unique")
        observed = native_host_capture(matching[0], hosts[host_id], complete=False)
        require(case["command"] == matching[0]["command"], "Failed case argv differs from actual host")
        if stage == "native_case":
            require(type(observed) is dict, "Failed native case lacks actual raw result")
            matching_case = [item for item in observed.get("cases", []) if item.get("id") == case.get("id")]
            require(len(matching_case) == 1 and all(case.get(key) == value for key, value in matching_case[0].items()),
                "Failed native case exit/streams/arguments differ from raw receipt")
            for item in case.get("pid_observations", []):
                if "receipt" in item:
                    raw = artifact(item["receipt"], available=item["receipt"]["available"])
                    if raw is not None:
                        require(json.loads(raw).get("pid") == item.get("pid"), "Failed authored readiness PID differs")


def expected_frame(message="authored process failure"):
    return {"protocol": "winbooksplit.result", "version": 1, "mode": "manual", "status": "invalid_input",
        "code": "invalid_start_pages", "message": message, "warnings": [], "fallback_modes": [],
        "exit_code": 2, "written_count": 0, "execution": None}


def frame_bytes(message="authored process failure"):
    return json.dumps(expected_frame(message), ensure_ascii=False, separators=(",", ":")).encode("utf-8")


def tail(data):
    data = data[-65536:]
    while data and data[0] & 0xC0 == 0x80:
        data = data[1:]
    return data.decode("utf-8")


def stream(record, field, data=None):
    text = record.get(field)
    require(type(text) is str and len(text.encode("utf-8")) <= 65536 and "\ufffd" not in text, field + ": unbounded or undecodable UTF-8 tail")
    count = record.get(field + "TotalBytes")
    require(type(count) is int and count >= 0 and type(record.get(field + "Truncated")) is bool,
            field + ": total/truncation required")
    require(record[field + "Truncated"] is (count > 65536), field + ": contradictory truncation")
    if data is not None:
        require(count == len(data) and text == tail(data), field + ": lost bytes, blank lines or final tail")
    elif count <= 65536:
        require(len(text.encode("utf-8")) == count, field + ": incomplete small stream")


def application_log_streams(text, stdout):
    """Separate the finalizer metadata from the two exact raw transport tails."""
    require(type(text) is str and type(stdout) is str, "Actual application log/stdout text required")
    parts = text.split("\r\n[OPERATION-OUTCOME] ")
    require(len(parts) == 2, "Exactly one provisional finalizer footer required")
    transport_log, footer = parts
    require(footer.endswith("\r\n") and "\r\n" not in footer[:-2], "Finalizer footer must be the terminal JSON line")
    provisional = json.loads(footer[:-2])
    final_lines = [line[len("[OUTCOME] "):] for line in stdout.splitlines() if line.startswith("[OUTCOME] ")]
    require(len(final_lines) == 1, "Exactly one actual final stdout outcome required")
    final = json.loads(final_lines[0])
    require(type(provisional) is dict and provisional == final, "Provisional finalizer/footer differs from actual final outcome")
    regions = transport_log.split("[STDOUT]\r\n", 1)
    require(len(regions) == 2, "Raw stdout section missing")
    regions = regions[1].split("\r\n[STDERR]\r\n", 1)
    require(len(regions) == 2, "Raw stderr section/separator missing")
    return regions[0], regions[1], provisional, final


def application_outcome(case):
    provisional, final = case.get("provisional_outcome"), case.get("final_outcome")
    require(type(provisional) is dict and type(final) is dict and provisional == final,
        "Application provisional/final outcome missing or contradicting")
    require(type(case.get("stdout")) is str, "Actual application stdout required")
    lines = [line[len("[OUTCOME] "):] for line in case["stdout"].splitlines() if line.startswith("[OUTCOME] ")]
    require(len(lines) == 1 and json.loads(lines[0]) == final, "Stored final outcome differs from actual stdout")
    engine = case.get("engine_record")
    require(final.get("protocol") == "winbooksplit.outcome" and type(final.get("version")) is int and final["version"] == 1
        and type(final.get("exit_code")) is int and final["exit_code"] == case.get("exit_code")
        and type(final.get("written_count")) is int and type(engine) is dict and final.get("engine_result") == engine
        and engine.get("protocol") == "winbooksplit.result" and type(engine.get("version")) is int and engine["version"] == 1
        and all(final.get(field) == engine.get(field) for field in ("status", "code", "exit_code", "mode")),
        "Application final native/engine outcome binding differs")
    success = case.get("kind") == "unicode-hostile"
    require((final.get("status"), final.get("code"), final.get("exit_code"), final["written_count"])
        == (("success", "split_complete", 0, 3) if success else ("invalid_input", "invalid_start_pages", 2, 0)),
        "Application final outcome does not match authored operation")
    execution = engine.get("execution")
    require(not success or type(execution) is dict, "Successful application lacks engine execution")
    require(final.get("final_directory") == (execution.get("final_directory") if success else None),
        "Application final publication differs from validated engine execution")


def native(case, host):
    flags(case, "passed", "actual_process", "unrelated_process_alive")
    require(case.get("id") == host + "-" + case.get("kind", ""), "Native ID differs from kind")
    result = case.get("process")
    require(type(result) is dict, "Missing actual supervisor result")
    flags(result, "ParentStopped", "DescendantsStopped", "StreamsComplete", "JobAssigned")
    for field in ("Pid", "ExitCode"):
        require(type(result.get(field)) is int, "Native integer status required")
    require(result["Pid"] > 0, "Native PID required")
    for field in ("StartError", "StreamError", "StopError"):
        require(field in result and result[field] is None, "Supervisor reported " + field)
    elapsed = result.get("ElapsedSeconds")
    require(type(elapsed) in (int, float) and math.isfinite(elapsed) and 0 <= elapsed < 10,
            "Native deadline/EOF was not bounded")
    kind = case["kind"]
    require(type(result.get("TimedOut")) is bool and type(result.get("Cancelled")) is bool, "Missing timeout/cancel labels")
    timed = kind in {"timeout-tree", "inherited-pipe", "detached-pipe"}
    require(result["TimedOut"] is timed and result["Cancelled"] is (kind == "cancel-tree"), "Wrong timeout/cancel outcome")
    data = {}
    status = {"flood": 2, "fast-tail": 23, "unicode-boundary": 0, "arguments": 0,
              "large-frame": 2, "oversized-frame": 2, "duplicate-frames": 2, "no-frame": 1,
              "malformed-frame": 1, "inherited-pipe": 0, "detached-pipe": 0}
    if kind in status:
        require(result["ExitCode"] == status[kind], "Native exit status changed")
    if kind == "flood":
        data = {"Stdout": b"O" * (2 * 1024 * 1024) + b"\n" + frame_bytes() + b"\n",
                "Stderr": b"E" * (2 * 1024 * 1024) + b"\nFINAL_STDERR_" + UNICODE.encode("utf-8")}
    elif kind == "fast-tail":
        data = {"Stdout": b"\n\nstdout blank\n\nno-newline-output", "Stderr": b"\n\nFINAL_STDERR_NO_NEWLINE"}
    elif kind == "unicode-boundary":
        data = {"Stdout": ("日" * (65536 // 3 + 1) + UNICODE).encode("utf-8"),
                "Stderr": ("ö" * (65536 // 2 + 1) + UNICODE).encode("utf-8")}
    elif kind == "large-frame":
        data = {"Stdout": frame_bytes("L" * 100000 + UNICODE), "Stderr": b""}
    elif kind == "oversized-frame":
        data = {"Stdout": frame_bytes("L" * 4096), "Stderr": b""}
    elif kind == "duplicate-frames":
        data = {"Stdout": frame_bytes() + b"\n" + frame_bytes() + b"\n", "Stderr": b""}
    elif kind == "no-frame":
        data = {"Stdout": b"Human diagnostic only", "Stderr": b""}
    elif kind == "malformed-frame":
        data = {"Stdout": b"{broken authored JSON", "Stderr": b""}
    for field in ("Stdout", "Stderr"):
        stream(result, field, data.get(field))
    records = result.get("ResultRecords")
    require(type(records) is list and all(type(value) is str for value in records), "ResultRecords array required")
    if kind in {"flood", "large-frame"}:
        require(len(records) == 1 and json.loads(records[0]) == expected_frame("L" * 100000 + UNICODE if kind == "large-frame" else "authored process failure"),
                "Structured frame lost behind bounded human tail")
        require(result.get("ResultError") is None and case.get("protocol_error") is None
                and case.get("parsed_result") == expected_frame("L" * 100000 + UNICODE if kind == "large-frame" else "authored process failure"),
                "Valid failure protocol was not preserved")
    if kind in {"oversized-frame", "duplicate-frames", "no-frame", "malformed-frame"}:
        require(case.get("parsed_result") is None and type(case.get("protocol_error")) is str and bool(case["protocol_error"]),
                "Invalid/absent/oversize/duplicate frame was accepted")
        if kind == "oversized-frame":
            require(bool(result.get("ResultError")), "Oversize frame omitted explicit bounded-capture error")
    if kind == "arguments":
        echo = json.loads(result["Stdout"])
        require(echo.get("argv") == ARGUMENTS and path(echo.get("python")) == path(case.get("python")), "Literal argv/interpreter changed")
        require(echo.get("encoding", "").lower().replace("-", "") == "utf8" and echo.get("utf8_env") == "1"
                and echo.get("io_env", "").lower() == "utf-8:backslashreplace", "Explicit child UTF-8 missing")
        require(case.get("arguments", [])[-len(ARGUMENTS):] == ARGUMENTS, "Receipt arguments differ from echo")
    if kind in {"timeout-tree", "cancel-tree", "inherited-pipe", "detached-pipe"}:
        observed = case.get("pid_observations")
        require(type(observed) is list and len(observed) == 3 and [item.get("role") for item in observed] == ["launched", "authored_parent", "authored_child"]
                and all(type(item.get("pid")) is int and item["pid"] > 0
                and item.get("running") is False for item in observed), "Owned parent/grandchild termination was not independently observed")
        require(observed[0]["pid"] == result["Pid"], "Observed original parent differs from supervisor PID")
        require(elapsed < 3, "Owned-tree stop exceeded bounded grace")


def validate_process_report(report, shell_paths, *, cleanup_complete=True):
    require(type(report) is dict and report.get("schema_version") == 1 and report.get("task_id") == "M2-T05"
            and report.get("result") == "PROCESS_REGRESSION_PASSED" and report.get("exit_code") == 0,
            "Wrong process receipt type/status")
    flags(report, "success", "source_unchanged", "baseline_guards_preserved", "input_and_neighbor_unchanged",
          "machine_settings_unchanged", "unrelated_process_removed")
    require(report.get("owned_temp_removed") is True if cleanup_complete else type(report.get("owned_temp_removed")) is bool,
            "Owned temporary cleanup proof missing")
    require(report.get("acceptance_ids") == ACCEPTANCE, "Exact acceptance IDs required")
    require(type(report.get("tested_path_sha256")) is dict and bool(report["tested_path_sha256"]), "Missing source digest")
    require({"engine/WinBookSplit.Process.ps1", "engine/WinBookSplit.Outcomes.json", "WinBookSplit.ps1", "WinBookSplit.bat",
        "tests/process/characterize_process.py", "tests/process/native_child.py", "tests/process/fake_engine.py",
        "tests/process/Probe-Process.ps1", "tests/process/validate_process_report.py"}.issubset(report["tested_path_sha256"]),
        "Promised supervisor/application/harness source hashes missing")
    for value in report["tested_path_sha256"].values():
        digest(value)
    unrelated = report.get("unrelated_process_observation")
    require(type(unrelated) is dict and unrelated.get("retained_popen_handle") is True
            and type(unrelated.get("retained_pid")) is int and unrelated["retained_pid"] > 0,
            "Unrelated retained actual process evidence missing")
    ready = unrelated.get("authored_ready")
    require(type(ready) is dict and ready.get("pid") == unrelated["retained_pid"]
            and path(ready.get("python")) == path(unrelated.get("executable")), "Unrelated self-PID/executable differs from retained process")
    for field in ("executable_sha256", "readiness_sha256"):
        digest(unrelated.get(field))
    require(unrelated.get("running_before_supervision") is True and unrelated.get("running_after_supervision") is True
            and unrelated.get("running_after_stop") is False and type(unrelated.get("retained_exit_code")) is int,
            "Unrelated fixture lacks independent alive-to-stopped observations")
    command = unrelated.get("command")
    require(type(command) is list and len(command) == 8 and path(command[0]) == path(unrelated["executable"])
            and command[1:6] == ["-I", "-B", "-X", "utf8", "-c"]
            and command[6] == "import json,os,pathlib,sys,time;pathlib.Path(sys.argv[1]).write_text(json.dumps({'pid':os.getpid(),'python':sys.executable}),encoding='utf-8');time.sleep(90)"
            and ntpath.basename(path(command[7])) == "unrelated-ready.json", "Unrelated actual argv proof differs")
    hosts = exact_cases(report.get("host_cases"), HOSTS)
    require(len(shell_paths) == 2 and {path(value) for value in shell_paths} == {path(host["shell_executable"]) for host in hosts.values()}, "Actual distinct promised hosts required")
    for identifier, host in hosts.items():
        flags(host, "passed")
        require(host.get("host_major") == (5 if identifier == "PS51" else 7)
                and type(host.get("host_version")) is str and host["host_version"].startswith("5.1." if identifier == "PS51" else "7."),
                "Supported actual host version missing")
        require(host.get("stored_policies") == host.get("policies_after") and type(host.get("stored_policies")) is list and host["stored_policies"], "Stored execution policies changed or missing")
        require(host.get("exit_code") == 0, "Actual process host failed")
    cases = exact_cases(report.get("native_cases"), [host + "-" + kind for host in HOSTS for kind in KINDS])
    captures = report.get("native_host_evidence")
    require(type(captures) is list and len(captures) == 2 and {item.get("host_id") for item in captures} == set(HOSTS),
        "Both complete raw native host receipts required")
    for capture in captures:
        host = hosts[capture["host_id"]]
        raw = native_host_capture(capture, host)
        raw_cases = exact_cases(raw.get("cases"), [host["id"] + "-" + kind for kind in KINDS])
        require(host.get("native_process") == capture["process"] and host.get("native_receipt_sha256") == capture["native_report"]["sha256"],
            "Raw native host evidence differs from stored host record")
        for identifier, raw_case in raw_cases.items():
            require(all(cases[identifier].get(key) == value for key, value in raw_case.items())
                and cases[identifier].get("command") == capture["command"] and path(cases[identifier].get("cwd")) == path(capture["cwd"]),
                "Native case status/streams/argv differ from the raw host receipt")
    for host in HOSTS:
        for kind in KINDS:
            native(cases[host + "-" + kind], host)
    apps = exact_cases(report.get("application_cases"), [host + "-app-" + kind for host in (*HOSTS, "BAT") for kind in ("flood", "fast-tail", "unicode-hostile")])
    for identifier, case in apps.items():
        application_outcome(case)
        flags(case, "passed", "actual_process", "input_unchanged", "neighbor_unchanged", "owned_outputs_removed", "literal_arguments_preserved", "log_utf8_roundtrip", "injection_marker_absent")
        markers = case.get("injection_marker_observations")
        require(type(markers) is list and [item.get("location") for item in markers] == ["wrapper_cwd", "engine_cwd", "book_directory"]
                and all(item.get("exists") is False and ntpath.basename(path(item.get("path"))) == "injected.txt" for item in markers),
                "Marker-bearing titles were not checked in all execution directories")
        host = identifier.split("-")[0]
        require(path(case.get("shell_executable")) == path(hosts["PS51" if host == "BAT" else host]["shell_executable"]), "App used another host")
        require(case.get("host_version") == hosts["PS51" if host == "BAT" else host]["host_version"], "App host version differs")
        require(type(case.get("command")) is list and case["command"] and type(case.get("stdout")) is str and type(case.get("stderr")) is str, "Actual native app evidence missing")
        require(case.get("batch_unchanged") is True if host == "BAT" else case.get("batch_unchanged") is None, "BAT preservation claim missing or misplaced")
        require(type(case.get("log_size_bytes")) is int and 0 < case["log_size_bytes"] < 400000, "Application log unbounded")
        digest(case.get("log_sha256"))
        console.validate_console_evidence(case.get("console_evidence"), case.get("final_outcome"), case.get("log_sha256"))
        summary = case.get("process_summary")
        require(type(summary) is dict, "Actual application supervisor summary missing")
        flags(summary, "ParentStopped", "DescendantsStopped", "StreamsComplete", "JobAssigned")
        require(summary.get("TimedOut") is False and summary.get("Cancelled") is False
                and type(summary.get("ExitCode")) is int and summary.get("ExitCode") == case.get("exit_code")
                and type(summary.get("Pid")) is int and summary["Pid"] > 0
                and type(summary.get("ResultRecordCount")) is int and summary.get("ResultRecordCount") == 1,
                "Application supervisor status contradicts child outcome")
        for field in ("StartError", "StreamError", "StopError", "ResultError"):
            require(field in summary and summary[field] is None, "Application supervisor failed: " + field)
        streams = {**summary, "Stdout": case.get("logged_stdout"), "Stderr": case.get("logged_stderr")}
        require(case.get("interaction", {}).get("engine_invocations") == [case.get("engine_invocation")], "Application interaction argv differs")
        for field in ("Stdout", "Stderr"):
            stream(streams, field)
        if case["kind"] in {"flood", "fast-tail"}:
            interactions.require_noninteractive(case.get("interaction"), [summary])
            require(case.get("stdin_utf8") == "" and ("-Mode Manual -StartPages '2,3' -NonInteractive -NoPause" in case.get("wrapper_text", "") if host != "BAT" else
                    case.get("controlled_modifications") == ["copied PS complete noninteractive terminal-frame choices"]),
                    "Fake terminal-frame control did not declare complete noninteractive choices")
            require(isinstance(case.get("wrapper_text"), str) and case.get("wrapper_sha256") == sha256(case["wrapper_text"].encode("utf-8")).hexdigest(),
                    "Actual noninteractive wrapper raw bytes/hash differ")
            require(case.get("exit_code") == 2 and case.get("written_count") == 0 and case.get("outputs") == []
                    and case.get("engine_record", {}).get("code") == "invalid_start_pages", "App lost native failure or announced output")
            receipt = case.get("child_receipt")
            require(type(receipt) is dict and receipt.get("argv") == case.get("expected_argv") and receipt.get("utf8_env") == "1"
                    and receipt.get("io_env") == "utf-8:backslashreplace", "Actual app child literal argv/encoding differs")
            flags(case, "final_stderr_preserved", "failure_outcome_preserved")
            if case["kind"] == "flood":
                expected_stdout = b"O" * (2 * 1024 * 1024) + b"\n" + frame_bytes() + b"\n"
                expected_stderr = b"E" * (2 * 1024 * 1024) + b"\nFINAL_STDERR_" + UNICODE.encode("utf-8")
            else:
                expected_stdout = b"\n\nstdout blank\n\n" + frame_bytes(UNICODE)
                expected_stderr = b"\n\nFINAL_STDERR_NO_NEWLINE_" + UNICODE.encode("utf-8")
            stream(streams, "Stdout", expected_stdout)
            stream(streams, "Stderr", expected_stderr)
        else:
            interactions.validate(case.get("interaction"), [summary], ["plan_ready"], ["execute"],
                                  execution=case.get("engine_record", {}).get("execution"))
            require(case.get("stdin_utf8") == "1\nY\nN\n\n" and case.get("controlled_modifications") == [],
                    "Real Unicode control did not explicitly confirm its plan")
            require(case.get("exit_code") == 0 and case.get("written_count") == 3 and len(case.get("outputs", [])) == 3, "Real Unicode PDF outputs missing")
            flags(case, "manifest_validated", "complete_page_content", "unicode_titles_preserved", "unicode_log_preserved", "source_readonly_observed")
            flags(case, "destination_bound_preview")
            require(case.get("page_ids") == [1, 2, 3] and case.get("source_content_sha256") == case.get("output_content_sha256")
                    and len(case["source_content_sha256"]) == 3, "Real physical page identity/content differs")
            require(case.get("expected_names") == [item.get("filename") for item in case["outputs"]], "Unicode/title filenames changed")
            titles = [UNICODE, "& % ! [literal] $(Set-Content injected.txt x)", "__import__('pathlib').Path('injected.txt').touch()"]
            require(case.get("source_titles") == titles and case.get("actual_titles") == titles, "Original Unicode/hostile title data changed")
            preview = case.get("expected_plan")
            require(type(preview) is dict and [entry.get("title") for entry in preview.get("entries", [])] == titles
                    and [[entry.get("start"), entry.get("end")] for entry in preview["entries"]] == [[0, 1], [1, 2], [2, 3]]
                    and [entry.get("filename") for entry in preview["entries"]] == case["expected_names"], "Bound preview does not match original Unicode titles/physical ranges/names")
            budget = case.get("independent_title_budget")
            require(type(budget) is int and 1 <= budget <= 50 and case["expected_names"] ==
                    [f"{number:02d} - {title[:budget].rstrip(' .')}.pdf" for number, title in enumerate(titles, 1)],
                    "Actual destination-aware Unicode names differ from independent title shortening")
    require(report.get("case_count") == len(cases) + len(apps), "Process case count differs")
