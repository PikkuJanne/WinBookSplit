"""Actual Windows PS5.1/PS7/BAT process, encoding and literal-data acceptance."""

from __future__ import annotations

import argparse
from copy import deepcopy
from datetime import datetime, timezone
from hashlib import sha256
import importlib.util
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time

from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[2]


def load(name, file):
    spec = importlib.util.spec_from_file_location(name, file)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


manual = load("wbs_process_manual", ROOT / "tests/manual/characterize_manual.py")
launchers = load("wbs_process_launchers", ROOT / "tests/manual/current_launchers.py")
paths = load("wbs_process_paths", ROOT / "tests/paths/characterize_paths.py")
conversion = load("wbs_process_content", ROOT / "tests/conversion/characterize_conversion.py")
runner = load("wbs_process_runner", ROOT / "tests/run_tests.py")
validator = load("wbs_process_validator", ROOT / "tests/process/validate_process_report.py")
history, require = manual.history, manual.require
APPLICATION = ("VERSION", "WinBookSplit.ps1", "WinBookSplit.bat", "requirements.txt",
    "engine/WinBookSplit.Paths.ps1", "engine/WinBookSplit.Diagnostics.ps1", "engine/WinBookSplit.Runtime.ps1",
    "engine/WinBookSplit.Process.ps1", "engine/WinBookSplit.Outcomes.json", "engine/winbooksplit_engine.py", "engine/winbooksplit_windows.py",
    "engine/winbooksplit_conversion.py", "engine/winbooksplit_job.py", "engine/WinBookSplit.Logging.ps1",
    "engine/WinBookSplit.Support.ps1", "Export-WinBookSplitDiagnostics.ps1")
OBSERVABILITY = {}


def owned_artifact(path):
    """Capture only this fixture's exact authored file, with a bounded raw copy."""
    result = {"path": str(path), "available": path.is_file() and not history.is_reparse(path)}
    if not result["available"]:
        return result
    details = path.stat()
    result.update(size_bytes=details.st_size, created_ns=details.st_ctime_ns, modified_ns=details.st_mtime_ns)
    if details.st_size > 16 * 1024 * 1024:
        result["capture_error"] = "Authored evidence exceeds the 16 MiB raw capture bound"
        return result
    raw = path.read_bytes()
    result.update(sha256=sha256(raw).hexdigest(), raw_text=raw.decode("utf-8", errors="strict"))
    return result


def attempt(stage, record):
    OBSERVABILITY.update(failure_stage=stage, last_case=record)


def failure_exit_code(error):
    return 130 if isinstance(error, KeyboardInterrupt) else \
        error.code if isinstance(error, SystemExit) and type(error.code) is int and error.code != 0 else 1


def failure_receipt(error, directory):
    source_error = None
    try:
        after = runner.source_manifest()
    except (OSError, ValueError) as source_failure:
        after, source_error = {}, str(source_failure)
    before = OBSERVABILITY.get("tested_path_sha256", {})
    return {"schema_version": 1, "task_id": "M2-T05", "result": "PROCESS_REGRESSION_FAILED", "success": False,
        "exit_code": failure_exit_code(error), "error": str(error), "error_type": type(error).__name__, "cleanup_safe": False,
        "workspace_retained": str(directory), "owned_temp_removed": not directory.exists(),
        "observed_at": datetime.now(timezone.utc).isoformat(),
        "tested_path_sha256": before, "source_after_sha256": after, "source_unchanged": bool(before) and before == after,
        "source_after_error": source_error,
        "failure_stage": OBSERVABILITY.get("failure_stage", "setup"), "last_case": OBSERVABILITY.get("last_case"),
        "host_cases": OBSERVABILITY.get("host_cases", []), "native_host_evidence": OBSERVABILITY.get("native_host_evidence", []),
        "completed_native_cases": OBSERVABILITY.get("completed_native_cases", []),
        "completed_application_cases": OBSERVABILITY.get("completed_application_cases", []),
        "unrelated_process_observation": OBSERVABILITY.get("unrelated_process_observation"),
        "scope": "Failed attempt evidence, never a process PASS or proof of the historical failure cause"}


def digest(path):
    return sha256(path.read_bytes()).hexdigest()


def running(pid):
    """Observe only explicit authored PIDs; no process-name scans or kills."""
    import ctypes
    from ctypes import wintypes
    api = ctypes.WinDLL("kernel32", use_last_error=True)
    api.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    api.OpenProcess.restype = wintypes.HANDLE
    api.GetExitCodeProcess.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
    api.GetExitCodeProcess.restype = wintypes.BOOL
    api.CloseHandle.argtypes = [wintypes.HANDLE]
    api.CloseHandle.restype = wintypes.BOOL
    handle = api.OpenProcess(0x1000, False, pid)
    if not handle:
        require(ctypes.get_last_error() == 87, "Cannot independently observe owned PID: " + str(pid))
        return False
    try:
        status = wintypes.DWORD()
        require(api.GetExitCodeProcess(handle, ctypes.byref(status)), "Owned PID status query failed")
        return status.value == 259
    finally:
        require(api.CloseHandle(handle), "Owned PID query handle did not close")


def host_probe(work):
    target = work / "host.ps1"
    target.write_text("[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)\n"
        "@{host_major=$PSVersionTable.PSVersion.Major;host_version=$PSVersionTable.PSVersion.ToString();syntax_error_count=0;"
        "stored_policies=@(Get-ExecutionPolicy -List|Where-Object {$_.Scope -ne 'Process'}|ForEach-Object {"
        "@{scope=$_.Scope.ToString();policy=$_.ExecutionPolicy.ToString()}})}|ConvertTo-Json -Depth 5\n", encoding="utf-8")
    return target


def native_cases(work, hosts, unrelated):
    records = []
    child = ROOT / "tests/process/native_child.py"
    for host in hosts:
        directory = work / (host["id"] + "-native")
        directory.mkdir()
        payload = {"helper": str(ROOT / "engine/WinBookSplit.Process.ps1"), "diagnostics": str(ROOT / "engine/WinBookSplit.Diagnostics.ps1"),
                   "python": sys.executable, "work": str(directory), "cases": []}
        for kind in validator.KINDS:
            arguments = ["-I", "-B", "-X", "utf8", str(child), kind]
            if kind == "arguments":
                arguments += validator.ARGUMENTS
            if kind in {"timeout-tree", "cancel-tree", "inherited-pipe", "detached-pipe"}:
                arguments += [str(directory / (kind + "-parent.json")), str(directory / (kind + "-child.json"))]
            payload["cases"].append({"id": host["id"] + "-" + kind, "kind": kind, "arguments": arguments,
                "timeout": 1.0 if kind in {"timeout-tree", "inherited-pipe", "detached-pipe"} else 8.0,
                "cancel": kind == "cancel-tree", "result_limit": 1024 if kind == "oversized-frame" else 8388608,
                "protocol": kind in {"flood", "large-frame", "oversized-frame", "duplicate-frames", "no-frame", "malformed-frame"}})
        payload_path, report_path = directory / "payload.json", directory / "native-report.json"
        payload_path.write_text(json.dumps(payload, ensure_ascii=True), encoding="utf-8")
        command = [host["shell_executable"], "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-File",
                   str(ROOT / "tests/process/Probe-Process.ps1"), "-PayloadPath", str(payload_path), "-ReportPath", str(report_path)]
        environment = history.clean_environment(directory)
        environment["PSModulePath"] = str(Path(host["shell_executable"]).parent / "Modules")
        capture = {"host_id": host["id"], "command": command, "cwd": str(directory),
            "payload": owned_artifact(payload_path), "process": None, "native_report": None}
        OBSERVABILITY.setdefault("native_host_evidence", []).append(capture)
        attempt("native_host", {"host_id": host["id"], "command": command, "cwd": str(directory), "passed": False})
        try:
            process = history.run_entrypoint(command, directory, environment=environment, stdin="", timeout=55)
        except history.EntryPointFailure as error:
            capture.update(native_report=owned_artifact(report_path), supervision_error={
                "error_type": type(error).__name__, "message": str(error), "tracked_parent_pid": error.pid,
                "raw_streams_available": False, "native_exit_available": False,
                "inner_tree_shutdown_proved": False})
            raise
        capture.update(process=process, native_report=owned_artifact(report_path))
        host.update(native_process=process)
        if process["exit_code"] != 0 or not report_path.is_file():
            raise history.EntryPointFailure("Native process host failed; inner owned-tree state unproved: " + repr(process), cleanup_safe=False, pid=None)
        try:
            observed = json.loads(report_path.read_text(encoding="utf-8"))
            for candidate in observed["cases"]:
                attempt("native_case", {**candidate, "passed": False, "command": command, "cwd": str(directory)})
                if any(candidate["process"].get(field) is not True for field in ("ParentStopped", "DescendantsStopped", "StreamsComplete")):
                    raise ValueError("Native owned-tree or stream stop unproved: " + candidate["id"])
        except (OSError, ValueError, TypeError, KeyError) as error:
            raise history.EntryPointFailure("Native receipt cannot prove tree/stream stop: " + str(error), cleanup_safe=False, pid=None) from error
        require(observed["host_version"] == host["host_version"] and observed["stored_policies"] == host["stored_policies"], "Native host evidence differs")
        for case in observed["cases"]:
            case.update(passed=False, python=sys.executable, command=command, host_version=host["host_version"], shell_executable=host["shell_executable"],
                        cwd=str(directory), unrelated_process_alive=unrelated.poll() is None and running(unrelated.pid))
            attempt("native_case", case)
            if case["kind"] in {"timeout-tree", "cancel-tree", "inherited-pipe", "detached-pipe"}:
                case["pid_observations"] = [{"role": "launched", "pid": case["process"]["Pid"], "running": running(case["process"]["Pid"])}]
                for suffix in ("parent", "child"):
                    target = directory / (case["kind"] + "-" + suffix + ".json")
                    captured = owned_artifact(target)
                    observation = {"role": "authored_" + suffix, "receipt": captured}
                    case["pid_observations"].append(observation)
                    if captured["available"] and "raw_text" in captured:
                        pid = json.loads(captured["raw_text"])["pid"]
                        require(type(pid) is int and pid > 0, "Authored PID receipt has no positive integer PID")
                        observation.update(pid=pid, running=running(pid), receipt_sha256=captured["sha256"])
                for item in case["pid_observations"][1:]:
                    require(item["receipt"]["available"] and "raw_text" in item["receipt"],
                        "Owned tree never reached authored PID control: " + item["receipt"]["path"])
            candidate = {**case, "passed": True}
            validator.native(candidate, host["id"])
            case["passed"] = True
            records.append(case)
            OBSERVABILITY.setdefault("completed_native_cases", []).append(deepcopy(case))
        host.update(native_process=process, native_receipt_sha256=digest(report_path))
    return records


def copy_application(app):
    app.mkdir()
    for name in APPLICATION:
        target = app / name
        target.parent.mkdir(exist_ok=True)
        shutil.copyfile(ROOT / name, target)
    return {name: digest(app / name) for name in APPLICATION}


def application_case(work, host, kind, generator, *, batch=False):
    label = ("BAT" if batch else host["id"]) + "-app-" + kind
    directory = work / label
    directory.mkdir()
    app, books, cwd, base = [directory / name for name in ("a", "b", "c", "o")]
    copied = copy_application(app)
    for target in (books, cwd, base):
        target.mkdir()
    source = books / "Janne ääkköset Größe 日本 [1] & %WBS_PROCESS_EXPAND%!WBS_PROCESS_EXPAND! (O'Neil).pdf"
    titles = [validator.UNICODE, "& % ! [literal] $(Set-Content injected.txt x)", "__import__('pathlib').Path('injected.txt').touch()"]
    source.write_bytes(generator._pdf_bytes({"title": "Original M2-T05 authored Unicode/literal PDF", "pages": 3,
        "outline": [{"title": title, "start_page": number, "children": []} for number, title in enumerate(titles, 1)]}))
    source_hash = digest(source)
    content = conversion.page_content(PdfReader(source))
    neighbor = books / "owned-neighbor.txt"
    neighbor.write_bytes(b"Preserve authored process neighbor\n")
    neighbor_hash = digest(neighbor)
    receipt_path = directory / "engine-receipt.json"
    injection_markers = [cwd / "injected.txt", app / "engine/injected.txt", books / "injected.txt"]
    if kind != "unicode-hostile":
        shutil.copyfile(ROOT / "tests/process/fake_engine.py", app / "engine/winbooksplit_engine.py")
    modifications = []
    if batch and kind != "unicode-hostile":
        # Shipped BAT accepts one literal source. This copied PS seam supplies
        # complete noninteractive choices for terminal-frame transport controls.
        text = (app / "WinBookSplit.ps1").read_text(encoding="utf-8-sig")
        seam = "# --- Validation ---"
        require(text.count(seam) == 1, "Copied noninteractive transport seam changed")
        text = text.replace(seam, "$Mode = 'Manual'; $StartPages = '2,3'; $NonInteractive = $true; $NoPause = $true\n"
            "$PSBoundParameters['Mode'] = $Mode; $PSBoundParameters['StartPages'] = $StartPages\n"
            "$PSBoundParameters['NonInteractive'] = $true; $PSBoundParameters['NoPause'] = $true\n" + seam)
        (app / "WinBookSplit.ps1").write_text(text, encoding="utf-8-sig")
        modifications.append("copied PS complete noninteractive terminal-frame choices")
    system = Path(os.environ["SystemRoot"])
    ps51 = system / "System32/WindowsPowerShell/v1.0/powershell.exe"
    environment = history.clean_environment(cwd)
    environment.update(PATH=os.pathsep.join((str(Path(sys.executable).parent), str(ps51.parent), str(system / "System32"), str(system))),
        PSModulePath=str(Path(host["shell_executable"]).parent / "Modules"), WBS_PROCESS_APP=str(app / "WinBookSplit.ps1"),
        WBS_PROCESS_BAT=str(app / "WinBookSplit.bat"), WBS_PROCESS_INPUT=str(source), WBS_PROCESS_BASE=str(base),
        WBS_PROCESS_PYTHON=sys.executable, WBS_PROCESS_EXPAND="UNEXPECTED_EXPANSION", WBS_PROCESS_KIND=kind,
        WBS_PROCESS_RECEIPT=str(receipt_path), WBS_PROCESS_CHILD=str(ROOT / "tests/process/native_child.py"))
    wrapper = directory / ("Invoke.cmd" if batch else "Invoke.ps1")
    if batch:
        wrapper.write_text('@echo off\nsetlocal DisableDelayedExpansion\n"%WBS_PROCESS_BAT%" "%WBS_PROCESS_INPUT%"\n', encoding="ascii", newline="\r\n")
        command = [str(system / "System32/cmd.exe"), "/d", "/v:off", "/c", str(wrapper)]
        observation = history.run([str(ps51), "-NoProfile", "-Command", "[Environment]::GetFolderPath('MyDocuments')"], cwd, environment=environment)
        require(observation["exit_code"] == 0 and observation["stdout"].strip(), "Actual BAT Documents observation failed")
        base = Path(observation["stdout"].strip()).resolve(strict=True)
    else:
        wrapper.write_text("[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)\n"
            "& $env:WBS_PROCESS_APP -InputFile $env:WBS_PROCESS_INPUT -OutputDirectory $env:WBS_PROCESS_BASE -PythonPath $env:WBS_PROCESS_PYTHON"
            + (" -Mode Manual -StartPages '2,3' -NonInteractive -NoPause" if kind != "unicode-hostile" else "") + "\n"
            "exit $LASTEXITCODE\n", encoding="utf-8")
        command = [host["shell_executable"], "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", str(wrapper)]
    expected_names = []
    preview = None
    if kind == "unicode-hostile":
        expected_engine = load("wbs_process_expected_engine_" + label.replace("-", "_"), ROOT / "engine/winbooksplit_engine.py")
        preview = manual.plain(expected_engine.preview_plan(expected_engine.prepare_split(source, "1", output_base=base)))
        require([entry["title"] for entry in preview["entries"]] == titles
                and [[entry["start"], entry["end"]] for entry in preview["entries"]] == [[0, 1], [1, 2], [2, 3]],
                "Destination-aware preview changed original title data or physical boundaries")
        naming = preview["output_naming"]
        def units(value):
            return len(str(value).encode("utf-16-le")) // 2
        directories = [base / (".WinBookSplit-stage-" + "0" * 32),
                       base / (naming["run_stem"] + "_20000101-000000_" + "0" * 32)]
        independent_budget = min(255, 259 - max(units(target) for target in directories) - 1)
        require(independent_budget == naming["filename_budget"], "Bound preview budget differs from independent complete Windows path measurement")
        title_budget = min(50, independent_budget - units("01 - .pdf"))
        require(title_budget > 0, "Authored Unicode/title fixture has no usable destination budget")
        expected_names = [f"{number:02d} - {title[:title_budget].rstrip(' .')}.pdf" for number, title in enumerate(titles, 1)]
        require(expected_names == [entry["filename"] for entry in preview["entries"]], "Bound preview filenames differ from independently shortened original Unicode titles")
    execution = log = owner = None
    members = None
    cleanup_safe = False
    record = None
    attempt("application_case", {"id": label, "kind": kind, "passed": False, "command": command,
        "cwd": str(cwd), "input_path": str(source), "output_base": str(base), "source_sha256_before": source_hash,
        "neighbor_sha256_before": neighbor_hash, "application_hashes": copied, "actual_process": None})
    try:
        try:
            with paths.read_only_source(source):
                read_only = bool(source.lstat().st_file_attributes & 1)
                process = history.run_entrypoint(command, cwd, environment=environment,
                    stdin="1\nY\nN\n\n" if kind == "unicode-hostile" else "", timeout=90)
            OBSERVABILITY["last_case"].update(process, actual_process=True)
        except history.EntryPointFailure as error:
            OBSERVABILITY["last_case"]["supervision_error"] = {
                "error_type": type(error).__name__, "message": str(error), "tracked_parent_pid": error.pid,
                "raw_streams_available": False, "native_exit_available": False, "inner_tree_shutdown_proved": False}
            raise
        log = launchers.exact_line(process["stdout"], "Log: ")
        require(log.parent.parent == base and log.name == "console.log" and not history.is_reparse(log.parent), "Actual console ownership escaped base")
        owner = json.loads((log.parent / ".WinBookSplit-console-owner.json").read_text(encoding="utf-8"))
        require(owner == {"run_id": log.parent.name.removeprefix(".WinBookSplit-console-"), "kind": "console"}, "Actual console ownership marker differs")
        raw = log.read_bytes()
        text = raw.decode("utf-8", errors="strict")
        OBSERVABILITY["last_case"].update(console_log=text, log_sha256=sha256(raw).hexdigest(),
            source_sha256_after=digest(source), neighbor_sha256_after=digest(neighbor))
        frames = [json.loads(line) for line in text.splitlines() if line.startswith('{')]
        require(len(frames) == 1, "Actual engine record missing or duplicated in console log")
        frame = frames[0]
        invocation = [json.loads(line.removeprefix("[ENGINE] ")) for line in text.splitlines() if line.startswith("[ENGINE] ")]
        require(len(invocation) == 1, "Actual engine invocation record missing")
        summaries = [json.loads(line.removeprefix("[PROCESS] ")) for line in text.splitlines() if line.startswith("[PROCESS] ")]
        require(len(summaries) == 1, "Actual supervisor summary missing")
        logged_stdout, logged_stderr, provisional_outcome, final_outcome = validator.application_log_streams(text, process["stdout"])
        for field, value in (("Stdout", logged_stdout), ("Stderr", logged_stderr)):
            if summaries[0][field + "Truncated"]:
                notice, separator, value = value.partition("\r\n")
                require(separator and notice.startswith("[TRUNCATED] "), "Missing bounded stream notice")
                if field == "Stdout":
                    logged_stdout = value
                else:
                    logged_stderr = value
        mode = "1" if kind == "unicode-hostile" else "manual"
        manual_data = "" if mode == "1" else "2,3"
        expected_argv = [str(source), str(base), mode, manual_data]
        interaction = launchers.interactions.capture(text)
        full_argv = ["-I", "-B", "-X", "utf8", str(app / "engine/winbooksplit_engine.py"), *expected_argv]
        if kind == "unicode-hostile":
            full_argv = launchers.interactions.validate(interaction, summaries, ["plan_ready"], ["execute"], execution=frame["execution"])
            require(full_argv[:9] == ["-I", "-B", "-X", "utf8", str(app / "engine/winbooksplit_engine.py"), *expected_argv]
                    and len(full_argv) == 11, "Real interactive engine literal argv differs")
        else:
            launchers.interactions.require_noninteractive(interaction, summaries)
        require(invocation[0]["arguments"] == full_argv, "Actual engine argv changed literal data")
        require(digest(source) == source_hash and digest(neighbor) == neighbor_hash and not any(target.exists() for target in injection_markers), "Actual app changed input/neighbor or evaluated injected data")
        record = {"id": label, "kind": kind, "passed": True, "actual_process": True,
            "shell_executable": host["shell_executable"], "host_version": host["host_version"], "command": command, "cwd": str(cwd),
            "batch_unchanged": digest(app / "WinBookSplit.bat") == copied["WinBookSplit.bat"] if batch else None,
            "application_hashes": copied, "engine_invocation": invocation[0], "expected_argv": expected_argv,
            "interaction": interaction, "controlled_modifications": modifications,
            "copied_ps_sha256": digest(app / "WinBookSplit.ps1"),
            "wrapper_text": wrapper.read_bytes().decode("utf-8"), "wrapper_sha256": digest(wrapper),
            "stdin_utf8": "1\nY\nN\n\n" if kind == "unicode-hostile" else "",
            "process_summary": summaries[0], "logged_stdout": logged_stdout, "logged_stderr": logged_stderr,
            "input_unchanged": True, "neighbor_unchanged": True, "literal_arguments_preserved": True,
            "log_utf8_roundtrip": text.encode("utf-8") == raw, "log_size_bytes": len(raw), "log_sha256": sha256(raw).hexdigest(),
            "injection_marker_absent": True, "injection_marker_observations": [
                {"location": location, "path": str(target), "exists": target.exists()}
                for location, target in zip(("wrapper_cwd", "engine_cwd", "book_directory"), injection_markers)],
            "engine_record": frame, "provisional_outcome": provisional_outcome, "final_outcome": final_outcome, **process}
        attempt("application_case", {**record, "passed": False, "console_log": text})
        record["console_evidence"] = launchers.authenticate_console_manifest(log, final_outcome,
            source_path=source, source_observation={"sha256": source_hash, "size_bytes": source.stat().st_size})
        require(len(raw) < 400000 and "\ufffd" not in text, "Actual application log unbounded or mojibake")
        if kind == "unicode-hostile":
            require(process["exit_code"] == 0 and frame["status"] == "success", "Actual Unicode/title split failed")
            execution = frame["execution"]
            outputs = manual.published_outputs(base, generator, execution)
            members = {".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", *(item["filename"] for item in outputs)}
            require([item["filename"] for item in outputs] == expected_names, "Actual hostile Unicode titles changed filenames")
            actual_content = [value for item in outputs for value in conversion.page_content(PdfReader(Path(execution["final_directory"]) / item["filename"]))]
            require(actual_content == content, "Actual Unicode PDF lost physical content")
            require(all("[Writing] " + name in text for name in expected_names), "Unicode log names failed roundtrip")
            record.update(outputs=outputs, written_count=len(outputs), expected_names=expected_names,
                manifest_validated=True, complete_page_content=True, unicode_titles_preserved=True, unicode_log_preserved=True,
                source_readonly_observed=read_only, page_ids=[value for item in outputs for value in item["page_ids"]],
                source_content_sha256=content, output_content_sha256=actual_content, destination_bound_preview=True,
                source_titles=titles, actual_titles=[entry["title"] for entry in execution["outputs"]], expected_plan=preview,
                independent_filename_budget=independent_budget, independent_title_budget=title_budget)
            require(record["actual_titles"] == titles and [entry["filename"] for entry in execution["outputs"]] == expected_names,
                    "Actual publication changed bound names or original Unicode/title data")
        else:
            require(process["exit_code"] == 2 and frame["code"] == "invalid_start_pages" and "Done." not in process["stdout"]
                    and "Output: " not in process["stdout"], "Controlled engine failure lost outcome through entrypoint")
            receipt = json.loads(receipt_path.read_text(encoding="utf-8"))
            require(receipt["argv"] == expected_argv, "Controlled actual engine argv differs")
            suffix = "FINAL_STDERR_" + validator.UNICODE if kind == "flood" else "FINAL_STDERR_NO_NEWLINE_" + validator.UNICODE
            require(suffix in text, "Actual engine final stderr tail lost")
            record.update(child_receipt=receipt, outputs=[], written_count=0, final_stderr_preserved=True, failure_outcome_preserved=True)
        cleanup_safe = True
    finally:
        if cleanup_safe:
            if execution is not None and members is not None:
                launchers.remove_known_directory(Path(execution["final_directory"]), base, members, ".WinBookSplit-owner.json",
                    {"schema_version": 1, "kind": "run", "run_id": execution["run_id"]})
            if log is not None and owner is not None:
                launchers.remove_console_directory(log, base, source_path=source,
                    source_observation={"sha256": source_hash, "size_bytes": source.stat().st_size})
    require(record is not None and not log.parent.exists() and (execution is None or not Path(execution["final_directory"]).exists()), "Owned application outputs remained")
    record["owned_outputs_removed"] = True
    OBSERVABILITY.setdefault("completed_application_cases", []).append(deepcopy(record))
    return record


def characterize(work, shells):
    require(os.name == "nt" and len(shells) == 2, "Actual Windows plus both supported hosts required")
    manual.trusted_original_sources()
    before = runner.source_manifest()
    OBSERVABILITY["tested_path_sha256"] = before
    probe = host_probe(work)
    hosts = [paths.host_observation(shell, work, probe) for shell in shells]
    OBSERVABILITY["host_cases"] = hosts
    require({host["id"] for host in hosts} == {"PS51", "PS7"}, "Actual distinct supported hosts required")
    # A venv executable may be a launcher shim. Use the actual ordinary pinned
    # base executable and verify authored os.getpid against the retained PID.
    unrelated_executable = (Path(sys.base_prefix) / "python.exe").resolve(strict=True)
    require(unrelated_executable.is_file() and not history.is_reparse(unrelated_executable), "Unrelated fixture needs the actual ordinary base interpreter")
    ready_path = work / "unrelated-ready.json"
    unrelated_code = "import json,os,pathlib,sys,time;pathlib.Path(sys.argv[1]).write_text(json.dumps({'pid':os.getpid(),'python':sys.executable}),encoding='utf-8');time.sleep(90)"
    unrelated_command = [str(unrelated_executable), "-I", "-B", "-X", "utf8", "-c", unrelated_code, str(ready_path)]
    unrelated = subprocess.Popen(unrelated_command,
        cwd=work, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    unrelated_record = {"command": unrelated_command, "executable": str(unrelated_executable), "executable_sha256": digest(unrelated_executable),
        "retained_pid": unrelated.pid, "retained_popen_handle": True}
    OBSERVABILITY["unrelated_process_observation"] = unrelated_record
    try:
        deadline = time.monotonic() + 5
        ready = None
        while time.monotonic() < deadline:
            if ready_path.is_file():
                try:
                    ready = json.loads(ready_path.read_text(encoding="utf-8"))
                    break
                except (OSError, ValueError):
                    pass  # Only the exact newly authored readiness file can be incomplete.
            require(unrelated.poll() is None, "Unrelated authored fixture exited before readiness")
            time.sleep(.01)
        require(type(ready) is dict and ready.get("pid") == unrelated.pid
                and Path(ready.get("python", "")).resolve(strict=True) == unrelated_executable,
                "Unrelated fixture did not prove actual self-PID/executable identity")
        unrelated_record.update(authored_ready=ready, readiness_sha256=digest(ready_path), running_before_supervision=running(unrelated.pid))
        require(unrelated_record["running_before_supervision"], "Unrelated fixture was not independently observed alive")
        native = native_cases(work, hosts, unrelated)
        unrelated_record["running_after_supervision"] = unrelated.poll() is None and running(unrelated.pid)
        require(unrelated_record["running_after_supervision"], "Supervisor killed the unrelated authored process")
    finally:
        if unrelated.poll() is None:
            unrelated.terminate()
        unrelated.wait(timeout=5)
        unrelated_record.update(retained_exit_code=unrelated.returncode, running_after_stop=running(unrelated.pid))
        require(unrelated_record["running_after_stop"] is False, "Tracked actual unrelated fixture remained alive after retained-handle stop")
    generator = history.load_generator()
    apps = []
    for host in hosts:
        for kind in ("flood", "fast-tail", "unicode-hostile"):
            apps.append(application_case(work, host, kind, generator))
    host51 = next(host for host in hosts if host["id"] == "PS51")
    for kind in ("flood", "fast-tail", "unicode-hostile"):
        apps.append(application_case(work, host51, kind, generator, batch=True))
    for host in hosts:
        after = paths.host_observation(Path(host["shell_executable"]), work, probe)
        require(after["stored_policies"] == host["stored_policies"], "Stored host policy changed")
        host["policies_after"] = after["stored_policies"]
    require(before == runner.source_manifest(), "Sources changed during process acceptance")
    return {"schema_version": 1, "task_id": "M2-T05", "result": "PROCESS_REGRESSION_PASSED", "success": True, "exit_code": 0,
        "observed_at": datetime.now(timezone.utc).isoformat(), "acceptance_ids": validator.ACCEPTANCE,
        "host_cases": hosts, "native_cases": native, "application_cases": apps, "case_count": len(native) + len(apps),
        "source_unchanged": True, "baseline_guards_preserved": True, "input_and_neighbor_unchanged": True,
        "machine_settings_unchanged": True, "unrelated_process_removed": unrelated.poll() is not None,
        "unrelated_process_observation": unrelated_record,
        "native_host_evidence": OBSERVABILITY.get("native_host_evidence", []),
        "tested_path_sha256": before, "immutable_original_commit": history.ORIGINAL_COMMIT,
        "environment": {"python": sys.version, "python_executable": sys.executable, "pypdf": history.pypdf.__version__},
        "not_run": ["human Ctrl+C/Explorer", "clean OS", "release package", "arbitrary unsupported encoding"],
        "limits": ["Timeout/cancel controls use a finite authored child and grandchild; cancellation is injected via a token.",
            "Native supervisor controls and controlled engine failure substitutes are distinguished from three real Unicode/title PDF splits.",
            "BAT wrapper supplies literal environment-variable data with delayed expansion disabled; original BAT bytes are unchanged.",
            "Only emitted authenticated application paths are removed; no private Documents enumeration or processing."]}


def main():
    OBSERVABILITY.clear()
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--shell-path", required=True, action="append", type=Path)
    args = parser.parse_args()
    try:
        require(sys.flags.isolated and sys.dont_write_bytecode, "Use the explicit supported developer Python -I -B")
        report_path = runner.new_external_path(args.report)
        require(len(args.shell_path) == 2 and len(set(args.shell_path)) == 2
                and all(path.is_absolute() and path.is_file() for path in args.shell_path), "Two actual absolute supported shell executables required")
        temporary = tempfile.TemporaryDirectory(prefix="P-")
        directory = Path(temporary.name).resolve()
        try:
            report = characterize(directory, [path.resolve() for path in args.shell_path])
            report["owned_temp_removed"] = False
            validator.validate_process_report(report, [str(path.resolve()) for path in args.shell_path], cleanup_complete=False)
            temporary.cleanup()
            report["owned_temp_removed"] = not directory.exists()
            validator.validate_process_report(report, [str(path.resolve()) for path in args.shell_path])
        except BaseException as error:
            temporary._finalizer.detach()
            failed = failure_receipt(error, directory)
            with report_path.open("x", encoding="utf-8", newline="\n") as stream:
                json.dump(failed, stream, indent=2, ensure_ascii=True)
                stream.write("\n")
            print("Failed process attempt; owned workspace retained: " + str(directory), file=sys.stderr)
            raise
        with report_path.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(report, stream, indent=2, ensure_ascii=True)
            stream.write("\n")
        print("Process acceptance passed: 26 actual supervisor controls and nine PS51/PS7/unchanged BAT application controls")
        return 0
    except BaseException as error:
        print("Process acceptance failed: " + str(error), file=sys.stderr)
        return failure_exit_code(error)


if __name__ == "__main__":
    raise SystemExit(main())
