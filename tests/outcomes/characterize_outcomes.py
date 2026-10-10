"""M3-T01 actual PS5.1/PS7/BAT final outcomes on authored local documents."""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
from hashlib import sha256
import importlib.util
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time

from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[2]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


process_tests = load("wbs_outcome_process", ROOT / "tests/process/characterize_process.py")
manual, history, require = process_tests.manual, process_tests.history, process_tests.require
launchers, paths, conversion = process_tests.launchers, process_tests.paths, process_tests.conversion
runner = process_tests.runner
validator = load("wbs_outcome_validator", ROOT / "tests/outcomes/validate_outcomes_report.py")
CLEANUP_SAFE = True
COMPLETED_CASES = []
LAST_CASE = None


def digest(path):
    return sha256(path.read_bytes()).hexdigest()


def source_identity(path):
    details = path.lstat()
    return {"sha256": digest(path), "size_bytes": details.st_size, "device": details.st_dev,
        "inode": details.st_ino, "attributes": details.st_file_attributes}


def anchored(text, before, after):
    require(text.count(before) == 1, "Copied application injection seam changed: " + before)
    return text.replace(before, after)


def controlled_application(app, kind, *, batch):
    original = process_tests.copy_application(app)
    text = (app / "WinBookSplit.ps1").read_text(encoding="utf-8-sig")
    modifications = []
    noninteractive = not kind.startswith("fallback-") and kind != "log-finalize"
    if batch:
        for parameter, value in (("OutputDirectory", "$env:WBS_OUTCOME_BASE"),
                ("PythonPath", "$env:WBS_OUTCOME_PYTHON"), ("CalibrePath", "$env:WBS_OUTCOME_CONVERTER")):
            text = anchored(text, "[string]$" + parameter + ",", "[string]$" + parameter + " = " + value + ",")
            modifications.append("copied parameter default " + parameter)
        if noninteractive:
            seam = "# --- Validation ---"
            text = anchored(text, seam, "$Mode = 'Manual'; $StartPages = '2,3'; $NonInteractive = $true; $NoPause = $true\n"
                "$PSBoundParameters['Mode'] = $Mode; $PSBoundParameters['StartPages'] = $StartPages\n"
                "$PSBoundParameters['NonInteractive'] = $true; $PSBoundParameters['NoPause'] = $true\n" + seam)
            modifications.append("copied PS complete noninteractive fault/deadline choices")
    if kind == "engine-timeout":
        text = anchored(text, "$ProcessTimeout = [Math]::Max(3600, $ConversionTimeout + 1800)", "$ProcessTimeout = 1")
        modifications.append("copied ProcessTimeout default 1 second")
    if kind == "conversion-timeout":
        text = anchored(text, "$ConversionTimeout = 1800", "$ConversionTimeout = 1")
        modifications.append("copied ConversionTimeout default 1 second")
    if kind in {"engine-cancel", "conversion-cancel"}:
        text = anchored(text, "    $script:lastProcessResult = Invoke-WinBookSplitProcess", 
            "    $authoredCancel = New-Object Threading.CancellationTokenSource\n"
            "    $authoredCancel.CancelAfter(1800)\n"
            "    $script:lastProcessResult = Invoke-WinBookSplitProcess")
        text = anchored(text, "-TimeoutSeconds $ProcessTimeout\n    }\n    $transport", "-TimeoutSeconds $ProcessTimeout -CancellationToken $authoredCancel.Token\n    $authoredCancel.Dispose()\n    }\n    $transport")
        modifications.append("copied supervisor call cancellation token after 1800ms")
    if kind == "log-finalize":
        start = text.index("function Complete-WinBookSplitConsoleLog")
        end = text.index("\n}", start) + 2
        function = text[start:end]
        declaration = function[:function.index("\n")]
        text = anchored(text, function, declaration + "\n    param([Parameter(Mandatory=$true)]$Outcome)\n"
            "    $script:consoleLogWriter.Dispose()\n    $script:consoleLogWriter = $null\n"
            "    throw 'AUTHORED_LOG_FINALIZE_FAILURE'\n}")
        modifications.append("copied console finalizer throws after closing actual log")
    if modifications:
        (app / "WinBookSplit.ps1").write_text(text, encoding="utf-8-sig")
    if kind in validator.ENGINE_FAULTS:
        (app / "engine/winbooksplit_engine.py").rename(app / "engine/saved_engine.py")
        shutil.copyfile(ROOT / "tests/outcomes/fault_engine.py", app / "engine/winbooksplit_engine.py")
        modifications.append("copied fault adapter loads actual saved engine")
    require(digest(app / "WinBookSplit.bat") == original["WinBookSplit.bat"], "Acceptance altered copied BAT bytes")
    actual = {name: digest(app / name) for name in process_tests.APPLICATION}
    if (app / "engine/saved_engine.py").exists():
        actual["engine/saved_engine.py"] = digest(app / "engine/saved_engine.py")
        require(actual["engine/saved_engine.py"] == original["engine/winbooksplit_engine.py"], "Saved engine differs from shipped bytes")
    return original, actual, modifications


def build_converter(work, ps51):
    executable = work / "OwnedConverter.exe"
    environment = history.clean_environment(work)
    source = ROOT / "tests/outcomes/OwnedConverter.cs"
    environment.update(WBS_OUTCOME_CONVERTER_SOURCE=str(source), WBS_OUTCOME_CONVERTER_EXE=str(executable))
    command = [str(ps51), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-Command",
        "Add-Type -Path $env:WBS_OUTCOME_CONVERTER_SOURCE -OutputAssembly $env:WBS_OUTCOME_CONVERTER_EXE -OutputType ConsoleApplication"]
    observed = history.run(command, work, environment=environment)
    require(observed["exit_code"] == 0 and executable.is_file(), "Authored native converter failed compilation: " + repr(observed))
    return executable, {**observed, "command": command, "source_sha256": digest(source), "executable_sha256": digest(executable)}


def marked_remove(directory, base, members, *, kind="run"):
    marker_path = directory / ".WinBookSplit-owner.json"
    marker = json.loads(marker_path.read_text(encoding="utf-8"))
    require(marker == {"schema_version": 1, "kind": kind, "run_id": marker["run_id"]}
        and re.fullmatch(r"[0-9a-f]{32}", marker["run_id"]), "Unexpected exact owned marker")
    launchers.remove_known_directory(directory, base, members, marker_path.name, marker)


def console_signal_entrypoint(command, cwd, environment, stdin, receipt, child_receipt):
    """Send an actual OS CTRL_C_EVENT only to this hidden private console."""
    import ctypes
    api = ctypes.WinDLL("kernel32", use_last_error=True)
    api.AttachConsole.argtypes = [ctypes.c_uint]
    api.AttachConsole.restype = ctypes.c_int
    api.FreeConsole.restype = ctypes.c_int
    api.SetConsoleCtrlHandler.argtypes = [ctypes.c_void_p, ctypes.c_int]
    api.SetConsoleCtrlHandler.restype = ctypes.c_int
    api.GenerateConsoleCtrlEvent.argtypes = [ctypes.c_uint, ctypes.c_uint]
    api.GenerateConsoleCtrlEvent.restype = ctypes.c_int
    api.CreateFileW.argtypes = [ctypes.c_wchar_p, ctypes.c_uint, ctypes.c_uint, ctypes.c_void_p, ctypes.c_uint, ctypes.c_uint, ctypes.c_void_p]
    api.CreateFileW.restype = ctypes.c_void_p
    api.CloseHandle.argtypes = [ctypes.c_void_p]
    api.CloseHandle.restype = ctypes.c_int
    class KeyEvent(ctypes.Structure):
        _fields_ = [("down", ctypes.c_int), ("repeat", ctypes.c_ushort), ("virtual_key", ctypes.c_ushort),
            ("scan", ctypes.c_ushort), ("character", ctypes.c_wchar), ("state", ctypes.c_uint)]
    class InputRecord(ctypes.Structure):
        _fields_ = [("event_type", ctypes.c_ushort), ("padding", ctypes.c_ushort), ("key", KeyEvent)]
    api.WriteConsoleInputW.argtypes = [ctypes.c_void_p, ctypes.POINTER(InputRecord), ctypes.c_uint, ctypes.POINTER(ctypes.c_uint)]
    api.WriteConsoleInputW.restype = ctypes.c_int
    startup = subprocess.STARTUPINFO()
    startup.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startup.wShowWindow = 0
    target = subprocess.Popen(command, cwd=cwd, env=environment, stdin=subprocess.PIPE,
        stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        creationflags=subprocess.CREATE_NEW_CONSOLE, startupinfo=startup)
    signal = {"actual_native_console_signal": True, "hidden_private_console": True,
        "target_pid": target.pid, "signal_sent": False, "forced_target_stop": False,
        "cmd_prompt_decision": None, "cmd_prompt_observed": False}
    chunks = {"stdout": [], "stderr": []}
    drain_errors = []
    def drain(stream, name):
        try:
            while value := os.read(stream.fileno(), 65536):
                chunks[name].append(value)
        except OSError as error:
            drain_errors.append(str(error))
    drains = [threading.Thread(target=drain, args=(stream, name), daemon=True)
        for stream, name in ((target.stdout, "stdout"), (target.stderr, "stderr"))]
    for thread in drains:
        thread.start()
    attached = False
    try:
        target.stdin.write(stdin.encode("utf-8"))
        target.stdin.flush()
        target.stdin.close()
        deadline = time.monotonic() + 12
        while (not receipt.exists() or not child_receipt.exists()) and time.monotonic() < deadline:
            require(target.poll() is None, "Ctrl+C target exited before authored child readiness")
            time.sleep(.01)
        require(receipt.exists() and child_receipt.exists(), "Ctrl+C tree never reached authored readiness")
        owned_pids = [json.loads(path.read_text(encoding="utf-8"))["pid"] for path in (receipt, child_receipt)]
        api.FreeConsole()
        require(api.AttachConsole(target.pid), "Cannot attach to actual private Ctrl+C target console")
        attached = True
        require(api.SetConsoleCtrlHandler(None, True), "Cannot protect only authored signal controller")
        signal["signal_sent"] = bool(api.GenerateConsoleCtrlEvent(0, 0))
        require(signal["signal_sent"], "Actual OS Ctrl+C signal failed")
        deadline = time.monotonic() + 15
        batch_target = Path(command[0]).name.lower() == "cmd.exe"
        while target.poll() is None and time.monotonic() < deadline:
            captured = b"".join(chunks["stdout"]).decode("utf-8", errors="replace")
            if batch_target and "Terminate batch job" in captured and signal["cmd_prompt_decision"] is None:
                result_rows = [json.loads(line.removeprefix("[OUTCOME] ")) for line in captured.splitlines() if line.startswith("[OUTCOME] ")]
                require(len(result_rows) == 1 and result_rows[0]["exit_code"] == 130
                    and all(not process_tests.running(pid) for pid in owned_pids), "CMD prompt preceded actual cancelled result/owned stop")
                handle = api.CreateFileW("CONIN$", 0xC0000000, 3, None, 3, 0, None)
                require(handle and handle != ctypes.c_void_p(-1).value, "Cannot open only attached private console input")
                try:
                    keys = (InputRecord * 4)()
                    for index, (character, virtual_key, down) in enumerate((("N", 0x4E, True), ("N", 0x4E, False), ("\r", 0x0D, True), ("\r", 0x0D, False))):
                        keys[index].event_type = 1
                        keys[index].key = KeyEvent(down, 1, virtual_key, 0, character, 0)
                    written = ctypes.c_uint()
                    require(api.WriteConsoleInputW(handle, keys, 4, ctypes.byref(written)) and written.value == 4,
                        "Private CMD prompt N/Enter input failed")
                finally:
                    require(api.CloseHandle(handle), "Private console input handle did not close")
                signal.update(cmd_prompt_decision="N then Enter", cmd_prompt_observed=True,
                    cmd_prompt_after_cancelled_outcome=True, cmd_prompt_owned_pids_stopped=True)
            time.sleep(.01)
        require(target.poll() is not None, "Actual private-console target did not exit after Ctrl+C/prompt")
        for thread in drains:
            thread.join(timeout=3)
        require(not drain_errors and all(not thread.is_alive() for thread in drains), "Actual Ctrl+C target streams did not drain safely")
        return {"exit_code": target.returncode, "stdout": b"".join(chunks["stdout"]).decode("utf-8", errors="replace"),
            "stderr": b"".join(chunks["stderr"]).decode("utf-8", errors="replace"), "timed_out": False,
            "console_signal": signal}
    except BaseException as error:
        if LAST_CASE is not None:
            LAST_CASE.update(console_signal=signal,
                stdout=b"".join(chunks["stdout"]).decode("utf-8", errors="replace"),
                stderr=b"".join(chunks["stderr"]).decode("utf-8", errors="replace"))
        raise history.EntryPointFailure("Actual OS Ctrl+C application stop was not proved: " + str(error), cleanup_safe=False, pid=target.pid) from error
    finally:
        if attached:
            # Ignore-Ctrl+C is inherited by later children, even when a child
            # receives a new console. Restore it before starting another case.
            restored = bool(api.SetConsoleCtrlHandler(None, False))
            signal["controller_ignore_restored"] = restored
            api.FreeConsole()
            require(restored, "Signal controller could not restore its inherited Ctrl+C behavior")
        if target.poll() is None:
            signal["forced_target_stop"] = True
            target.terminate()
            target.wait(timeout=5)
        for thread in drains:
            thread.join(timeout=1)
        target.stdout.close()
        target.stderr.close()


def case(work, host, kind, generator, converter, unrelated, *, batch=False):
    global CLEANUP_SAFE, LAST_CASE
    label = ("BAT" if batch else host["id"]) + "-" + kind
    directory = work / label
    directory.mkdir()
    app, books, cwd, base = [directory / name for name in ("a", "b", "c", "o")]
    copied, actual, modifications = controlled_application(app, kind, batch=batch)
    for target in (books, cwd, base):
        target.mkdir()
    ebook = kind in {"converter-failure", "conversion-cancel", "conversion-timeout", "conversion-ctrlc"}
    source = books / ("Authored ää [1] book.epub" if ebook else "Authored ää [1] book.pdf")
    if ebook:
        # Converter fault tests operate on authored bytes and never certify a
        # real Calibre conversion; real EPUB/AZW3 remain inherited full gates.
        source.write_bytes(b"Authored converter failure input only\n")
    elif kind == "original-failure":
        source.write_bytes(b"Authored invalid PDF bytes\n")
    else:
        source.write_bytes(generator._pdf_bytes({"title": "Original M3-T01 flat PDF", "pages": 3, "outline": []}))
    neighbor, prior = books / "neighbor.txt", base / "prior-output.pdf"
    neighbor.write_bytes(b"Authored unrelated source neighbor\n")
    prior.write_bytes(b"Authored earlier output stays byte identical\n")
    originals = {str(path): source_identity(path) for path in (source, neighbor, prior)}
    receipt, child_receipt = directory / "fault.json", directory / "child.json"
    candidate_python = directory / "missing-python.exe" if kind == "startup-dependency" else Path(sys.executable)
    system = Path(os.environ["SystemRoot"])
    environment = history.clean_environment(cwd)
    environment.update(PATH=os.pathsep.join((str(Path(sys.executable).parent), str(Path(host["shell_executable"]).parent), str(system / "System32"), str(system))),
        PSModulePath=str(Path(host["shell_executable"]).parent / "Modules"),
        WBS_OUTCOME_APP=str(app / "WinBookSplit.ps1"), WBS_OUTCOME_BAT=str(app / "WinBookSplit.bat"),
        WBS_OUTCOME_INPUT=str(source), WBS_OUTCOME_BASE=str(base), WBS_OUTCOME_PYTHON=str(candidate_python),
        WBS_OUTCOME_CONVERTER=str(converter), WBS_OUTCOME_KIND=kind, WBS_OUTCOME_RECEIPT=str(receipt),
        WBS_OUTCOME_CHILD_RECEIPT=str(child_receipt))
    wrapper = directory / ("Invoke.cmd" if batch else "Invoke.ps1")
    noninteractive = not kind.startswith("fallback-") and kind != "log-finalize"
    if batch:
        wrapper.write_text('@echo off\nsetlocal DisableDelayedExpansion\n"%WBS_OUTCOME_BAT%" "%WBS_OUTCOME_INPUT%"\n', encoding="ascii", newline="\r\n")
        command = [str(system / "System32/cmd.exe"), "/d", "/v:off", "/c", str(wrapper)]
    else:
        wrapper.write_text("[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)\n"
            "& $env:WBS_OUTCOME_APP -InputFile $env:WBS_OUTCOME_INPUT -OutputDirectory $env:WBS_OUTCOME_BASE "
            "-PythonPath $env:WBS_OUTCOME_PYTHON -CalibrePath $env:WBS_OUTCOME_CONVERTER"
            + (" -Mode Manual -StartPages '2,3' -NonInteractive -NoPause" if noninteractive else "") + "\nexit $LASTEXITCODE\n", encoding="utf-8")
        command = [host["shell_executable"], "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", str(wrapper)]
    stdin = "" if noninteractive else "M\n2,3\nY\n"
    if kind.startswith("fallback-"):
        stdin = "1\n" + {"fallback-success": "M\n2,3\nY\nN\n\n", "fallback-failure": "M\n2,no\n", "fallback-cancel": "C\n"}[kind]
    CLEANUP_SAFE = False
    LAST_CASE = {"id": label, "command": command, "cwd": str(cwd), "output_base": str(base)}
    if kind.endswith("ctrlc"):
        observed = console_signal_entrypoint(command, cwd, environment, stdin, receipt, child_receipt)
    else:
        observed = history.run_entrypoint(command, cwd, environment=environment, stdin=stdin, timeout=60)
    observed["timed_out"] = False
    LAST_CASE.update(observed)
    outcome_rows = [json.loads(line.removeprefix("[OUTCOME] ")) for line in observed["stdout"].splitlines() if line.startswith("[OUTCOME] ")]
    require(len(outcome_rows) == 1, "Exactly one final stdout outcome required for " + label + ": " + repr(observed))
    outcome = outcome_rows[0]
    require(outcome["exit_code"] == observed["exit_code"], "Final stdout/native outcome mismatch")
    log_lines = [line.removeprefix("Log: ").strip() for line in observed["stdout"].splitlines() if line.startswith("Log: ")]
    log = Path(log_lines[0]) if log_lines else None
    frames, transports = [], []
    if kind == "startup-dependency":
        require(log is None and {path.name for path in base.iterdir()} == {prior.name}, "Dependency failure touched output storage")
        log_text = ""
    else:
        require(len(log_lines) == 1 and log.parent.parent == base and log.name == "console.log", "Output log escaped authored base")
        log_text = log.read_text(encoding="utf-8-sig")
        frames = [json.loads(line) for line in log_text.splitlines() if line.startswith('{')]
        transports = [json.loads(line.removeprefix("[PROCESS] ")) for line in log_text.splitlines() if line.startswith("[PROCESS] ")]
        require(transports and all(item["ParentStopped"] and item["DescendantsStopped"] and item["StreamsComplete"] for item in transports),
            "Application did not prove engine tree/drains stopped")
    fault = json.loads(receipt.read_text(encoding="utf-8")) if receipt.exists() else None
    pid_observations = []
    for value in ([fault] if fault else []) + ([json.loads(child_receipt.read_text(encoding="utf-8"))] if child_receipt.exists() else []):
        pid_observations.append({"pid": value["pid"], "running": process_tests.running(value["pid"])})
    require(all(value["running"] is False for value in pid_observations), "Authored operation descendant remained alive")
    CLEANUP_SAFE = True
    require(observed["exit_code"] == validator.EXPECTED[kind], "Final native outcome differs for " + label + ": " + repr(observed))
    if kind in validator.INTERRUPTED:
        require(fault is not None and child_receipt.exists() and len(pid_observations) == 2, "Interrupted operation never reached actual authored tree")
    require(unrelated.poll() is None and process_tests.running(unrelated.pid), "Cancelled/timed-out run stopped unrelated Python")
    after_operation = {str(path): source_identity(path) for path in (source, neighbor, prior)}
    require(originals == after_operation, "Run changed source/neighbor/prior output")
    success = kind == "fallback-success"
    require(("Done." in observed["stdout"]) == success and ("Output: " in observed["stdout"]) == success,
        "Completion message disagrees with final operation")
    if kind.startswith("fallback-"):
        require(len(frames) == len(transports) == 1 and frames[0] == outcome["engine_result"], "Fallback did not finish one captured engine session")
    final_result = outcome.get("engine_result")
    execution = final_result.get("execution") if isinstance(final_result, dict) else None
    publication = None
    content = actual_content = publication_manifest = publication_manifest_sha = None
    owned_children = {prior.name}
    if execution is not None:
        require(kind in {"fallback-success", "log-finalize"}, "Failed operation published chapters")
        publication = manual.published_outputs(base, generator, execution)
        for item in publication:
            target = Path(execution["final_directory"]) / item["filename"]
            item.update(page_count=len(PdfReader(target).pages), size_bytes=target.stat().st_size)
        content = conversion.page_content(PdfReader(source))
        actual_content = [value for item in publication for value in conversion.page_content(PdfReader(Path(execution["final_directory"]) / item["filename"]))]
        require(content == actual_content and [value for item in publication for value in item["page_ids"]] == [1, 2, 3], "Published PDF omitted/changed physical pages")
        manifest_path = Path(execution["final_directory"]) / "WinBookSplit_Manifest.json"
        publication_manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        publication_manifest_sha = digest(manifest_path)
        owned_children.add(Path(execution["final_directory"]).name)
    diagnostic = final_result.get("diagnostic") if isinstance(final_result, dict) else None
    diagnostic_path = Path(diagnostic["record_path"]).parent if diagnostic and diagnostic.get("record_path") else None
    diagnostic_record = None
    diagnostic_hash = None
    stage = None
    if fault and fault.get("stage"):
        stage = Path(fault["stage"])
    elif fault and fault.get("output"):
        stage = Path(fault["output"]).parent
    retained = stage is not None and stage.exists()
    if diagnostic_path:
        diagnostic_record = json.loads((diagnostic_path / "failure.json").read_text(encoding="utf-8"))
        diagnostic_hash = digest(diagnostic_path / "failure.json")
        require(diagnostic_path.parent == base and diagnostic_record["status"] in {"failed", "cancelled", "timeout"}, "Failure diagnostics invalid")
        owned_children.add(diagnostic_path.name)
    if kind in {"output-write", "manifest-finalize", "publish-rename", "cleanup-refusal", "converter-failure", "conversion-timeout"}:
        require(diagnostic_record is not None, "Ordinary failure omitted recoverable diagnostic record")
    if kind in {"converter-failure", "conversion-timeout"}:
        observed_converter = diagnostic.get("conversion")
        require(observed_converter and "AUTHORED_CONVERTER_FINAL_STDERR" in observed_converter["stderr_tail"]
            and observed_converter["argv"] == [str(converter), str(source), fault["output"], "--output-profile", "tablet"], "Actual converter status/argv/final stderr missing")
        if kind == "converter-failure":
            require(observed_converter["exit_code"] == 17, "Actual converter failure exit was overwritten")
    if retained:
        require(stage.parent == base and re.fullmatch(r"\.WinBookSplit-stage-[0-9a-f]{32}", stage.name), "Retained interrupted stage escaped output base")
        owned_children.add(stage.name)
        if kind == "cleanup-refusal":
            require((stage / "authored-unrelated.txt").read_bytes() == b"Keep this unregistered acceptance member\n"
                and diagnostic["cleanup_complete"] is False and diagnostic["cleanup_error"], "Cleanup refusal hid/deleted unrelated member")
        elif kind not in {"engine-cancel", "engine-timeout", "conversion-cancel", "engine-ctrlc", "conversion-ctrlc"}:
            raise RuntimeError("Ordinary failure did not clean known owned stage: " + label)
    if log:
        owned_children.add(log.parent.name)
    require({path.name for path in base.iterdir()} == owned_children, "Failure left undisclosed/final-looking/unexpected output")
    record = {"id": label, "kind": kind, "passed": True, "actual_process": True,
        "shell_executable": host["shell_executable"], "host_version": host["host_version"], "batch": batch,
        "command": command, "cwd": str(cwd), "output_base": str(base), "input_path": str(source), "outcome": outcome,
        "engine_records": frames, "process_summaries": transports, "application_source_sha256": copied,
        "controlled_application_sha256": actual, "controlled_modifications": modifications,
        "batch_bytes_unchanged": digest(app / "WinBookSplit.bat") == copied["WinBookSplit.bat"],
        "source_neighbor_prior_unchanged": True, "source_observations": originals, "after_operation_source_observations": after_operation,
        "unrelated_process_alive": True, "unrelated_pid": unrelated.pid, "fault_receipt": fault,
        "pid_observations": pid_observations, "retained_staging": str(stage) if retained else None,
        "known_stage_cleanup_complete": not retained, "final_publication": publication,
        "source_page_content_sha256": content, "output_page_content_sha256": actual_content,
        "publication_manifest": publication_manifest, "publication_manifest_sha256": publication_manifest_sha,
        "failure_diagnostic_record": diagnostic_record, "failure_diagnostic_sha256": diagnostic_hash,
        "converter_path": str(converter) if ebook else None, "converter_sha256": digest(converter) if ebook else None,
        "final_publication_validated": publication is not None, "final_completion_truthful": True,
        "log_sha256": digest(log) if log else None, "log_finalizer_failure_visible": "AUTHORED_LOG_FINALIZE_FAILURE" in observed["stdout"],
        "output_members_before_cleanup": sorted(owned_children), "stdin_utf8": stdin,
        "wrapper_text": wrapper.read_bytes().decode("utf-8"), "wrapper_sha256": digest(wrapper),
        "interaction": launchers.interactions.capture(log_text), **observed}
    record["console_evidence"] = launchers.authenticate_console_manifest(log, outcome,
        source_path=source, source_observation=originals[str(source)], allow_unfinalized=kind == "log-finalize") if log else None
    if kind != "startup-dependency":
        if noninteractive:
            launchers.interactions.require_noninteractive(record["interaction"], transports)
        else:
            stages = ["no_plan"] if kind == "fallback-cancel" else ["no_plan", "input_ready"] if kind == "fallback-failure" else \
                ["no_plan", "input_ready", "plan_ready"] if kind == "fallback-success" else ["input_ready", "plan_ready"]
            actions = ["cancel"] if kind == "fallback-cancel" else ["retry", "starts"] if kind == "fallback-failure" else \
                ["retry", "starts", "execute"] if kind == "fallback-success" else ["starts", "execute"]
            launchers.interactions.validate(record["interaction"], transports, stages, actions,
                starts="2,no" if kind == "fallback-failure" else "2,3", execution=execution)
    if execution:
        marked_remove(Path(execution["final_directory"]), base, {".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", *(entry["filename"] for entry in execution["outputs"])})
    if diagnostic_path:
        marked_remove(diagnostic_path, base, {".WinBookSplit-owner.json", "failure.json"}, kind="failed")
    if retained:
        if fault.get("stage"):
            require([stage.stat().st_dev, stage.stat().st_ino] == fault["stage_identity"], "Retained fault stage identity changed")
            require({item.name for item in stage.iterdir()} == set(fault["members"]), "Retained stage introduced unknown members")
            for name, value in fault["members"].items():
                item = stage / name
                require([item.stat().st_dev, item.stat().st_ino] == value["identity"] and digest(item) == value["sha256"], "Retained stage member changed")
            members = set(fault["members"])
        else:
            members = {".WinBookSplit-owner.json", Path(fault["output"]).name}
        marked_remove(stage, base, members)
    if log:
        launchers.remove_console_directory(log, base, outcome, source_path=source,
            source_observation=originals[str(source)], allow_unfinalized=kind == "log-finalize")
    after_cleanup = {str(path): source_identity(path) for path in (source, neighbor, prior)}
    require({item.name for item in base.iterdir()} == {prior.name} and originals == after_cleanup, "Harness cleanup reached unrelated file")
    record.update(owned_application_outputs_removed=True, post_cleanup_source_neighbor_prior_unchanged=True,
        post_cleanup_source_observations=after_cleanup, output_members_after_cleanup=[prior.name])
    validator.validate_case(record)
    COMPLETED_CASES.append(record)
    LAST_CASE = record
    return record


def characterize(work, shells):
    require(os.name == "nt" and len(shells) == 2, "Actual Windows and both explicit supported hosts required")
    manual.trusted_original_sources()
    source_before = runner.source_manifest()
    probe = process_tests.host_probe(work)
    hosts = [paths.host_observation(shell, work, probe) for shell in shells]
    require({host["id"] for host in hosts} == {"PS51", "PS7"}, "Both actual distinct supported hosts required")
    host51 = next(host for host in hosts if host["id"] == "PS51")
    converter, compilation = build_converter(work, Path(host51["shell_executable"]))
    generator = history.load_generator()
    executable = (Path(sys.base_prefix) / "python.exe").resolve(strict=True)
    ready = work / "unrelated.json"
    command = [str(executable), "-I", "-B", "-X", "utf8", "-c",
        "import json,os,pathlib,sys,time;pathlib.Path(sys.argv[1]).write_text(json.dumps({'pid':os.getpid(),'python':sys.executable}),encoding='utf-8');time.sleep(900)", str(ready)]
    unrelated = subprocess.Popen(command, cwd=work, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    cases = []
    try:
        deadline = time.monotonic() + 5
        while not ready.exists() and time.monotonic() < deadline:
            require(unrelated.poll() is None, "Authored unrelated process exited before readiness")
            time.sleep(.01)
        identity = json.loads(ready.read_text(encoding="utf-8"))
        require(identity["pid"] == unrelated.pid and Path(identity["python"]).resolve() == executable
            and process_tests.running(unrelated.pid), "Unrelated process did not prove actual PID/interpreter identity")
        for host in hosts:
            for kind in validator.EXPECTED:
                cases.append(case(work, host, kind, generator, converter, unrelated))
        for kind in validator.EXPECTED:
            cases.append(case(work, host51, kind, generator, converter, unrelated, batch=True))
        alive_after = unrelated.poll() is None and process_tests.running(unrelated.pid)
        require(alive_after, "Operation stopped unrelated process")
    finally:
        if unrelated.poll() is None:
            unrelated.terminate()
        unrelated.wait(timeout=5)
    require(not process_tests.running(unrelated.pid), "Tracked unrelated fixture remained alive after retained-handle termination")
    for host in hosts:
        after = paths.host_observation(Path(host["shell_executable"]), work, probe)
        host["policies_after"] = after["stored_policies"]
        require(host["stored_policies"] == host["policies_after"], "Stored shell execution policy changed")
    require(source_before == runner.source_manifest(), "Sources changed during acceptance")
    return {"schema_version": 1, "task_id": "M3-T01", "result": "OUTCOME_REGRESSION_PASSED", "success": True, "exit_code": 0,
        "observed_at": datetime.now(timezone.utc).isoformat(), "acceptance_ids": validator.ACCEPTANCE,
        "host_cases": hosts, "cases": cases, "case_count": len(cases), "source_unchanged": True,
        "tested_path_sha256": source_before, "converter_compilation": compilation,
        "unrelated_process_observation": {"command": command, "executable_sha256": digest(executable), "actual_identity": identity,
            "retained_pid": unrelated.pid, "alive_after_all_operations": alive_after, "running_after_retained_handle_stop": process_tests.running(unrelated.pid)},
        "machine_settings_unchanged": True, "baseline_guards_preserved": True,
        "environment": {"python": sys.version, "python_executable": sys.executable, "pypdf": history.pypdf.__version__},
        "not_run": ["human Ctrl+C", "Explorer drag/drop", "release package", "real Calibre in this layer"],
        "limits": ["Cancellation is an injected supervisor token in an owned copied application; no human Ctrl+C claim.",
            "Authored native converter probes/failures do not certify actual EPUB/AZW3 conversion; full retains real Calibre gates.",
            "Interrupted outer engine stages are retained by the application; harness recovery uses exact authenticated synthetic members.",
            "BAT bytes equal the shipped BAT; copied PowerShell defaults provide controlled external paths and fault deadlines."]}


def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--shell-path", required=True, action="append", type=Path)
    args = parser.parse_args()
    try:
        require(sys.flags.isolated and sys.dont_write_bytecode, "Use explicit supported developer Python -I -B")
        target = runner.new_external_path(args.report)
        require(len(args.shell_path) == 2 and len(set(args.shell_path)) == 2 and all(path.is_absolute() and path.is_file() for path in args.shell_path), "Both actual supported shell paths required")
        temporary = tempfile.TemporaryDirectory(prefix="O-")
        work = Path(temporary.name).resolve()
        try:
            report = characterize(work, [path.resolve() for path in args.shell_path])
        except history.EntryPointFailure as error:
            if not error.cleanup_safe:
                temporary._finalizer.detach()
                print("Unproved owned tree stop; preserve authored workspace: " + str(work), file=sys.stderr)
            else:
                temporary.cleanup()
            raise
        except BaseException:
            if CLEANUP_SAFE:
                temporary.cleanup()
            else:
                temporary._finalizer.detach()
                print("Unproved operation stop; preserve authored workspace: " + str(work), file=sys.stderr)
            raise
        else:
            temporary.cleanup()
        report["owned_temp_removed"] = not work.exists()
        validator.validate_outcomes_report(report, [str(path.resolve()) for path in args.shell_path])
        with target.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(report, stream, indent=2, ensure_ascii=True)
            stream.write("\n")
        print("Outcomes acceptance passed: 51 actual operation/fallback/interrupt/finalization controls including six OS Ctrl+C events")
        return 0
    except (OSError, RuntimeError, ValueError, subprocess.TimeoutExpired) as error:
        if 'target' in locals() and not target.exists():
            with target.open("x", encoding="utf-8", newline="\n") as stream:
                json.dump({"schema_version": 1, "task_id": "M3-T01", "result": "OUTCOME_REGRESSION_FAILED",
                    "success": False, "exit_code": 1, "error": str(error), "completed_cases": COMPLETED_CASES,
                    "last_case": LAST_CASE, "cleanup_safe": CLEANUP_SAFE,
                    "owned_temp_removed": not work.exists() if 'work' in locals() else None}, stream, indent=2, ensure_ascii=True)
                stream.write("\n")
        print("Outcomes acceptance failed: " + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
