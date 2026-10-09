"""M2-T02 actual Windows literal paths, safe names and destination budgets."""

from __future__ import annotations

import argparse
from contextlib import contextmanager, redirect_stdout
from datetime import datetime, timezone
import errno
from io import StringIO
import json
import os
from pathlib import Path
import platform
import re
import shutil
import stat
import subprocess
import sys
import tempfile
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[2]
ENGINE = ROOT / "engine/winbooksplit_engine.py"


def load(name, path):
    import importlib.util
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


manual = load("wbs_paths_manual", ROOT / "tests/manual/characterize_manual.py")
launchers = load("wbs_paths_launchers", ROOT / "tests/manual/current_launchers.py")
plans = load("wbs_paths_plans", ROOT / "tests/plans/characterize_plan.py")
runner = load("wbs_paths_runner", ROOT / "tests/run_tests.py")
history, require = manual.history, manual.require

TITLE_CASES = [
    ("forbidden-control", 'A<>:"/\\|?*\x01\x7fB', "01 - AB.pdf"),
    ("trailing-dots", "Ending...   ", "02 - Ending.pdf"),
    ("reserved-title", "CON", "03 - CON.pdf"),
    ("punctuation-only", "... !!!", "04 - Section 4.pdf"),
    ("unicode", "日本語 åäö 😀", "05 - 日本語 åäö 😀.pdf"),
    ("whitespace", "  spaced \t \n words  ", "06 - spaced words.pdf"),
    ("case-first", "Same", "07 - Same.pdf"),
    ("case-second", "same", "08 - same.pdf"),
    ("long-astral", "Å😀" * 40, "09 - " + "Å😀" * 16 + "Å.pdf"),
]


@contextmanager
def exclusive_native(path):
    """Actual owned sharing-denial control; no ACL or policy edits."""
    windows = load("wbs_paths_native_lock", ROOT / "engine/winbooksplit_windows.py")
    api = windows._windows_api()
    handle = api.CreateFileW(str(path), 0x80000000, 0, None, 3, 0x02000000 | 0x00200000, None)
    import ctypes
    require(handle is not None and handle != ctypes.c_void_p(-1).value,
            "Cannot acquire the authored exclusive Windows sharing control")
    try:
        yield
    finally:
        require(bool(api.CloseHandle(handle)), "Authored native lock did not close")


@contextmanager
def read_only_source(path):
    """Change and restore only this exact owned source's read-only attribute."""
    original = path.lstat()
    identity = original.st_dev, original.st_ino
    require(stat.S_ISREG(original.st_mode) and not history.is_reparse(path), "Read-only control needs an ordinary owned source")
    os.chmod(path, stat.S_IREAD)
    try:
        require(path.lstat().st_file_attributes & stat.FILE_ATTRIBUTE_READONLY, "Actual Windows read-only source attribute was not set")
        yield
    finally:
        observed = path.lstat()
        require((observed.st_dev, observed.st_ino) == identity and not history.is_reparse(path),
                "Refuse to restore attributes on a replaced source")
        os.chmod(path, original.st_mode)
        require(path.lstat().st_file_attributes == original.st_file_attributes, "Owned source attributes were not restored")


def make_pdf(path, generator, pages, outline):
    with path.open("xb") as stream:
        stream.write(generator._pdf_bytes({"pages": pages, "title": "Original M2-T02 synthetic path fixture",
                                           "outline": outline}))
    require(generator.page_ids(path) == list(range(1, pages + 1)), "Generated page identities differ")
    return path


def long_base(parent, units):
    """Create ordinary owned components with an exact conservative path length."""
    require(runner.path_units(str(parent)) < units, "External TEMP prefix exceeds the authored test budget")
    current = parent
    while runner.path_units(str(current)) < units:
        remaining = units - runner.path_units(str(current)) - 1
        require(remaining > 0, "Cannot construct the exact authored path budget")
        component = min(40, remaining)
        if remaining - component == 1:
            component -= 1  # A later component needs one separator and a name.
        current /= "b" * component
        current.mkdir()
    require(runner.path_units(str(current)) == units, "Long base measurement differs")
    return current


def remove_run(base, execution):
    final = manual.published_directory(base, execution)
    launchers.remove_known_directory(final, base,
        {".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", *(entry["filename"] for entry in execution["outputs"])},
        ".WinBookSplit-owner.json", {"schema_version": 1, "kind": "run", "run_id": execution["run_id"]})


def parity_record(source, base, prepared, execution, outputs):
    preview = manual.plain(prepared.plan)
    content = plans.page_content(source)
    final = manual.published_directory(base, execution)
    actual_content = [value for entry in outputs for value in plans.page_content(final / entry["filename"])]
    record = {"passed": True, "mode": preview["mode"], "total_pages": preview["total_pages"],
              "preview_entries": preview["entries"], "coverage": preview["coverage"], "outputs": outputs,
              "source_identity": preview["source_identity"], "writer_result": execution,
              "written_count": len(outputs), "neighbor_unchanged": True, "no_replanning_or_reopening": True,
              "original_page_content_sha256": content, "page_content_sha256": actual_content,
              "manifest_validated": True, "input_unchanged": True, "input_resolved": str(source.resolve()),
              "input_sha256": history.file_digest(source), "output_base": str(base.resolve()),
              "final_path_utf16_units": [runner.path_units(str(final / entry["filename"])) for entry in outputs],
              "stage_path_utf16_units": [runner.path_units(str(base / (".WinBookSplit-stage-" + execution["run_id"]) / entry["filename"]))
                                          for entry in outputs]}
    runner.validate_plan_parity(record)
    runner.validate_plan_page_content(record)
    return record


def write_case(engine, source, base, generator, mode="1", manual_data=None):
    neighbor = base / "synthetic-neighbor.txt"
    neighbor.write_bytes(b"Do not alter this authored neighboring output\n")
    source_hash = history.file_digest(source)
    prepared = engine.prepare_split(source, mode, manual_data, output_base=base)
    preview = manual.plain(prepared.plan)
    reader = engine.PdfReader
    source_reader_calls, planner_calls = [], []

    def guarded_reader(path, *args, **kwargs):
        if isinstance(path, (str, os.PathLike)) and os.path.normcase(os.path.realpath(path)) == os.path.normcase(str(source.resolve())):
            source_reader_calls.append(str(path))
            raise RuntimeError("Execution reopened the original source instead of its captured reader")
        return reader(path, *args, **kwargs)

    def planner(*args, **kwargs):
        planner_calls.append(True)
        raise RuntimeError("Execution replanned or prepared the immutable destination-bound plan")

    with patch.object(engine, "PdfReader", side_effect=guarded_reader), \
            patch.object(engine, "prepare_split", side_effect=planner), \
            patch.object(engine, "plan_manual_starts", side_effect=planner), \
            patch.object(engine, "plan_level1", side_effect=planner), \
            patch.object(engine, "plan_level2", side_effect=planner), redirect_stdout(StringIO()):
        execution = manual.plain(engine.execute_split(prepared, base))
    outputs = manual.published_outputs(base, generator, execution)
    manual.check_base_members(base, execution, [neighbor.name])
    record = parity_record(source, base, prepared, execution, outputs)
    require(history.file_digest(source) == source_hash and neighbor.read_bytes() == b"Do not alter this authored neighboring output\n",
            "Path writer changed an input or neighbor")
    require(manual.plain(prepared.plan) == preview, "Execution changed its destination-bound immutable preview")
    remove_run(base, execution)
    record.update(owned_outputs_removed=not Path(execution["final_directory"]).exists(), preview_unchanged=True)
    record.update(no_replanning_observation="direct-api-planner-and-source-reopen-traps",
                  planner_calls_during_execute=len(planner_calls), source_reader_calls_during_execute=len(source_reader_calls))
    require(record["owned_outputs_removed"], "Held cleanup left the published owned path")
    return record


def host_observation(shell, work, probe):
    environment = history.clean_environment(work)
    environment.update(WBS_PATHS_ROOT=str(ROOT), PSModulePath=str(shell.parent / "Modules"))
    command = [str(shell), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-File", str(probe)]
    result = history.run(command, work, environment=environment)
    require(result["exit_code"] == 0, "Actual path host syntax/version observation failed: " + repr(result))
    observation = json.loads(result["stdout"])
    require(observation["host_major"] in (5, 7) and observation["syntax_error_count"] == 0,
            "Actual supported host or production syntax invalid")
    return {"id": "PS51" if observation["host_major"] == 5 else "PS7", "passed": True,
            "shell_executable": str(shell), "command": command, "cwd": str(work), **result, **observation}


def launcher_case(identifier, shell, source, base, cwd, generator, engine, wrapper, *, batch=False,
                  input_argument=None, expected_success=False, locked=False):
    system = Path(os.environ["SystemRoot"])
    ps51 = system / "System32/WindowsPowerShell/v1.0/powershell.exe"
    environment = history.clean_environment(cwd)
    environment.update(PATH=os.pathsep.join((str(Path(sys.executable).parent), str(ps51.parent), str(system / "System32"), str(system))),
                       PSModulePath=str(shell.parent / "Modules"), WBS_PATHS_SCRIPT=str(ROOT / "WinBookSplit.ps1"),
                       WBS_PATHS_INPUT=input_argument if input_argument is not None else str(source),
                       WBS_PATHS_BASE=str(base), WBS_PATHS_EXPAND="UNEXPECTED_EXPANSION", WBS_PATHS_BAT=str(ROOT / "WinBookSplit.bat"))
    if identifier.endswith("provider-pdf"):
        environment["WBS_PATHS_PROVIDER.pdf"] = "Authored non-filesystem item"
    command = [str(system / "System32/cmd.exe"), "/d", "/v:off", "/c", str(wrapper)] if batch else \
        [str(shell), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", str(wrapper)]
    neighbor = None if batch else base / "synthetic-neighbor.txt"
    if neighbor is not None:
        neighbor.write_bytes(b"Only this external owned base may be used\n")
    source_hash = history.file_digest(source) if source.is_file() else None
    final, execution, log, marker = None, None, None, None
    cleanup_safe, run_validated = True, False
    process = None
    try:
        try:
            if locked:
                with exclusive_native(source):
                    try:
                        source.read_bytes()
                    except OSError:
                        lock_verified = True
                    else:
                        raise RuntimeError("Actual exclusive source lock did not prevent a reader")
                    process = history.run_entrypoint(command, cwd, environment=environment, stdin="M\n4,7\n\n", timeout=60)
            else:
                lock_verified = False
                process = history.run_entrypoint(command, cwd, environment=environment, stdin="M\n4,7\n\n", timeout=60)
        except history.EntryPointFailure as error:
            cleanup_safe = error.cleanup_safe
            raise
        log_lines = [line for line in process["stdout"].splitlines() if line.startswith("Log: ")]
        output_lines = [line for line in process["stdout"].splitlines() if line.startswith("Output: ")]
        payloads = []
        if log_lines:
            log = launchers.exact_line(process["stdout"], "Log: ")
            require(log.is_absolute() and log.name == "console.log" and log.parent.parent == base
                    and log.parent.resolve(strict=True) == log.parent and not history.is_reparse(log.parent)
                    and re.fullmatch(r"\.WinBookSplit-console-[0-9a-f]{32}", log.parent.name), "Path console location escaped its explicit base")
            marker = json.loads((log.parent / ".WinBookSplit-console-owner.json").read_text(encoding="utf-8"))
            require(marker == {"run_id": log.parent.name.removeprefix(".WinBookSplit-console-"), "kind": "console"}, "Path console owner differs")
            payloads = [json.loads(line) for line in log.read_text(encoding="utf-8").splitlines() if line.startswith("{")]
        expected_exit = 0 if expected_success else 6 if identifier.endswith("corrupt-pdf") else 2
        require(process["exit_code"] == expected_exit, "Actual path launcher exit differs: " + identifier + ": " + repr(process))
        if expected_success:
            require(len(payloads) == 1 and payloads[0]["status"] == "success" and payloads[0]["exit_code"] == 0,
                    "Actual literal source did not produce exactly one successful engine record")
            execution = payloads[0]["execution"]
            final = launchers.exact_line(process["stdout"], "Output: ")
            require(str(final) == execution["final_directory"] and execution["source_identity"]["path"] == str(source.resolve())
                    and execution["source_identity"]["sha256"] == source_hash, "Launcher wrote a wildcard match or changed the literal source")
            outputs = manual.published_outputs(base, generator, execution)
            run_validated = True
            prepared = engine.prepare_split(source, "manual", "4,7", output_base=base)
            record = parity_record(source, base, prepared, execution, outputs)
            sizes = [line.removeprefix("Size:").strip() for line in process["stdout"].splitlines() if line.startswith("Size:")]
            expected_size = f"{source.stat().st_size / (1024 * 1024):.2f} MB"
            require(len(sizes) == 1 and sizes[0].replace(",", ".") == expected_size, "Metadata came from a wildcard-decoy input")
            require("%WBS_PATHS_EXPAND%!WBS_PATHS_EXPAND!" in process["stdout"], "A launcher expanded literal percent/exclamation text")
            record.update(input_size_display=expected_size, reported_size_display=sizes[0], metadata_size_matches=True,
                          literal_expansion_preserved=True, input_argument=environment["WBS_PATHS_INPUT"], engine_record=payloads[0])
        else:
            require(not output_lines and "Done." not in process["stdout"] and "[Writing]" not in process["stdout"],
                    "Rejected input announced output or success")
            corrupt = identifier.endswith("corrupt-pdf")
            require(bool(payloads) is corrupt, "File/provider validation failed to reject before engine execution")
            if corrupt:
                require(len(payloads) == 1 and payloads[0]["code"] == "unreadable_document" and payloads[0]["execution"] is None
                        and payloads[0]["written_count"] == 0, "Parser did not reject the corrupt actual file")
            record = {"passed": True, "outputs": [], "successful_final_count": 0, "engine_called": corrupt,
                      "error_code": payloads[0]["code"] if corrupt else "literal_preflight_rejected",
                      "native_lock_verified": lock_verified}
        require(source_hash is None or history.file_digest(source) == source_hash, "Launcher changed the literal input")
        require(neighbor is None or neighbor.read_bytes() == b"Only this external owned base may be used\n", "Launcher changed an output-base neighbor")
        record.update(id=identifier, shell_executable=str(shell), actual_process=True, command=command, cwd=str(cwd),
                      input_unchanged=True, neighbor_unchanged=True, **process)
        if expected_success:
            record.update(source_read_only_attribute_observed=bool(source.lstat().st_file_attributes & stat.FILE_ATTRIBUTE_READONLY),
                          no_replanning_observation="supported-by-direct-api-and-shared-plan-controls")
    finally:
        if cleanup_safe:
            if run_validated:
                remove_run(base, execution)
            if log is not None and marker is not None:
                launchers.remove_known_directory(log.parent, base, {".WinBookSplit-console-owner.json", "console.log"},
                                                 ".WinBookSplit-console-owner.json", marker)
    require(final is None or not final.exists(), "Owned literal final remains")
    require(log is None or not log.parent.exists(), "Owned literal console remains")
    record["owned_outputs_removed"] = True
    if not batch:
        require({path.name for path in base.iterdir()} == {neighbor.name}, "Literal launcher left unexpected base members")
    return record


def actual_launchers(work, shells, fixtures, generator, engine, inputs):
    probe = work / "Observe-PathHost.ps1"
    probe.write_text("[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)\n"
        "$total=0; $checked=@('WinBookSplit.ps1','engine/WinBookSplit.Paths.ps1')\n"
        "foreach($name in $checked){$tokens=$null;$errors=$null;"
        "[Management.Automation.Language.Parser]::ParseFile((Join-Path $env:WBS_PATHS_ROOT $name),[ref]$tokens,[ref]$errors)|Out-Null;$total+=$errors.Count}\n"
        "$policies=@(Get-ExecutionPolicy -List|Where-Object {$_.Scope -ne 'Process'}|ForEach-Object {"
        "[pscustomobject]@{scope=$_.Scope.ToString();policy=$_.ExecutionPolicy.ToString()}})\n"
        "[pscustomobject]@{host_major=$PSVersionTable.PSVersion.Major;host_version=$PSVersionTable.PSVersion.ToString();"
        "syntax_checked=$checked;syntax_error_count=$total;stored_policies=$policies}|ConvertTo-Json -Depth 6 -Compress\n", encoding="utf-8")
    hosts = [host_observation(shell, work, probe) for shell in shells]
    require({host["id"] for host in hosts} == {"PS51", "PS7"}, "Both actual supported hosts are mandatory")
    by_host = {host["id"]: Path(host["shell_executable"]) for host in hosts}
    wrapper = work / "Invoke-LiteralPaths.ps1"
    wrapper.write_text("[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)\n"
                       "& $env:WBS_PATHS_SCRIPT -InputFile $env:WBS_PATHS_INPUT -OutputDirectory $env:WBS_PATHS_BASE\n"
                       "exit $LASTEXITCODE\n", encoding="utf-8")
    batch_wrapper = work / "Invoke-LiteralPaths.cmd"
    batch_wrapper.write_text('@echo off\nsetlocal DisableDelayedExpansion\n"%WBS_PATHS_BAT%" "%WBS_PATHS_INPUT%"\n', encoding="ascii", newline="\r\n")
    source_dir = work / "literal inputs [dir] & (group)"
    source_dir.mkdir()
    source = source_dir / "book [1] O'Neil & (notes) %WBS_PATHS_EXPAND%!WBS_PATHS_EXPAND! 章节.PdF"
    source.write_bytes(fixtures["simple10"].read_bytes() + b"\n% Authored metadata size distinction\n" + b" " * 100000)
    decoy = source_dir / source.name.replace("[1]", "1")
    decoy.write_bytes(fixtures["nested12"].read_bytes())
    inputs.update({str(source): history.file_digest(source), str(decoy): history.file_digest(decoy)})
    successes, rejections = [], []
    for host_id in ("PS51", "PS7"):
        cwd = work / (host_id + "-unrelated-cwd")
        cwd.mkdir()
        base = work / (host_id + " output [base] O'Neil & (group) %! 章节")
        base.mkdir()
        argument = os.path.relpath(source, cwd)
        with read_only_source(source):
            record = launcher_case(host_id + "-literal", by_host[host_id], source, base, cwd, generator, engine, wrapper,
                                   input_argument=argument, expected_success=True)
        record["source_attributes_restored"] = True
        record["wildcard_decoy_unchanged"] = history.file_digest(decoy) == inputs[str(decoy)]
        successes.append(record)
        for kind in ("directory-pdf", "provider-pdf", "corrupt-pdf", "unreadable-held-file"):
            rejected_base = work / (host_id + "-" + kind + "-base")
            rejected_base.mkdir()
            rejected_source = cwd / (kind + ".pdf")
            argument = None
            if kind == "directory-pdf":
                rejected_source.mkdir()
            elif kind == "provider-pdf":
                argument = "Env:WBS_PATHS_PROVIDER.pdf"
            else:
                rejected_source.write_bytes(b"Authored non-PDF bytes\n" if kind == "corrupt-pdf" else fixtures["simple10"].read_bytes())
                inputs[str(rejected_source)] = history.file_digest(rejected_source)
            rejections.append(launcher_case(host_id + "-" + kind, by_host[host_id], rejected_source, rejected_base, cwd,
                                           generator, engine, wrapper, input_argument=argument, locked=kind == "unreadable-held-file"))
    environment = history.clean_environment(work)
    environment["PSModulePath"] = str(by_host["PS51"].parent / "Modules")
    observed = history.run([str(by_host["PS51"]), "-NoProfile", "-Command", "[Environment]::GetFolderPath('MyDocuments')"],
                           work, environment=environment)
    require(observed["exit_code"] == 0 and observed["stdout"].strip(), "Actual BAT default output-base observation failed")
    documents = Path(observed["stdout"].strip()).resolve(strict=True)
    require(documents.is_dir() and not history.is_reparse(documents), "BAT default base must be an ordinary actual directory")
    cwd = work / "BAT-unrelated-cwd"
    cwd.mkdir()
    with read_only_source(source):
        record = launcher_case("BAT-literal", by_host["PS51"], source, documents, cwd, generator, engine, batch_wrapper,
                               input_argument=str(source), expected_success=True, batch=True)
    record["source_attributes_restored"] = True
    record["wildcard_decoy_unchanged"] = history.file_digest(decoy) == inputs[str(decoy)]
    successes.append(record)
    for host in hosts:
        after = host_observation(Path(host["shell_executable"]), work, probe)
        require(after["stored_policies"] == host["stored_policies"], "Stored machine/user execution policies changed")
        host["policies_after"] = after["stored_policies"]
    return hosts, successes, rejections


def filenames(engine, work, generator, inputs):
    source = make_pdf(work / "safe-titles.pdf", generator, len(TITLE_CASES),
                      [{"title": title, "start_page": index, "children": []}
                       for index, (_, title, _) in enumerate(TITLE_CASES, 1)])
    inputs[str(source)] = history.file_digest(source)
    base = work / "safe-title-base"
    base.mkdir()
    record = write_case(engine, source, base, generator)
    require([item["filename"] for item in record["outputs"]] == [expected for _, _, expected in TITLE_CASES],
            "Written names differ from independent safe-title expectations")
    record["title_cases"] = [{"id": identifier, "passed": True, "original_title": title, "filename": entry["filename"]}
                             for (identifier, title, _), entry in zip(TITLE_CASES, record["outputs"], strict=True)]
    return record


def chapters(engine, work, generator, inputs):
    records = []
    for mode in ("manual", "1", "2"):
        nodes = [{"title": "Chapter " + str(index), "start_page": index, "children": []} for index in range(1, 121)]
        outline = [{"title": "Parent", "start_page": 1, "children": nodes}] if mode == "2" else nodes
        source = make_pdf(work / ("chapters120-mode-" + mode + ".pdf"), generator, 120, outline)
        inputs[str(source)] = history.file_digest(source)
        base = work / ("chapters120-" + mode + "-base")
        base.mkdir()
        record = write_case(engine, source, base, generator, mode,
                            ",".join(map(str, range(1, 121))) if mode == "manual" else None)
        names = [item["filename"] for item in record["outputs"]]
        require(len(names) == 120 and names == sorted(names)
                and all(name.startswith(f"{index:03d} - ") for index, name in enumerate(names, 1)), "120-section ordering or number width differs")
        record.update(section_count=120, number_width=3)
        records.append(record)
    return records


def destinations(engine, work, generator, inputs):
    source = make_pdf(work / ("LongBook" + "s" * 35 + ".pdf"), generator, 3,
                      [{"title": "章" * 100, "start_page": 1, "children": []}])
    inputs[str(source)] = history.file_digest(source)
    long_parent = work / "long-bound"
    long_parent.mkdir()
    base = long_base(long_parent, 175)
    prepared = engine.prepare_split(source, "1", output_base=base)
    record = write_case(engine, source, base, generator)
    record["output_naming"] = manual.plain(prepared.plan["output_naming"])
    record["shortened"] = record["outputs"][0]["filename"] != "01 - " + "章" * 50 + ".pdf"
    require(record["shortened"], "Authored long destination did not exercise full-path-aware shortening")
    failures = []
    over_parent = work / "over-budget"
    over_parent.mkdir()
    too_long = long_base(over_parent, 190)
    for identifier, code in runner.PATH_DESTINATION_FAILURE_CODES.items():
        target = too_long if identifier in ("over-budget", "unbound-long-base") else work / (identifier + "-base")
        if not target.exists():
            target.mkdir()
        neighbor = target / (identifier + "-neighbor.txt")
        neighbor.write_bytes(b"Preserve authored failure neighbor\n")
        diagnostic, calls, lock_verified = None, [], False
        source_hash = history.file_digest(source)
        if identifier in ("unbound-long-base", "changed-bound-base"):
            prepared = engine.prepare_split(source, "1", output_base=base if identifier == "changed-bound-base" else None)
            try:
                with redirect_stdout(StringIO()):
                    engine.execute_split(prepared, target)
            except engine.PlanError as error:
                actual_code, message = error.code, str(error)
            else:
                raise RuntimeError("Frozen destination or impossible unbound budget was accepted")
            method = "Actual execute_split rejects the prepared binding before allocation"
        else:
            writer = engine.write_slice

            def full_writer(*args, **kwargs):
                calls.append(str(args[3]))
                if len(calls) == 2:
                    raise OSError(errno.ENOSPC, "Authored full-target simulation after one exclusive slice")
                writer(*args, **kwargs)

            if identifier == "unwritable-held-base":
                with exclusive_native(target), redirect_stdout(StringIO()):
                    result = manual.plain(engine.run_split(source, target, "manual", "1,2"))
                    lock_verified = True
                method = "Actual Windows held exclusive base prevents production guard access; no ACL edits"
            elif identifier == "full-target":
                with patch.object(engine, "write_slice", side_effect=full_writer), redirect_stdout(StringIO()):
                    result = manual.plain(engine.run_split(source, target, "manual", "1,2"))
                method = "Explicit ENOSPC writer injection after one real exclusive written slice"
            else:
                with redirect_stdout(StringIO()):
                    result = manual.plain(engine.run_split(source, target, "1"))
                method = "Actual run_split rejects all created-path budgets before allocation"
            expected_exit = 6 if identifier == "full-target" else 2
            require(result["exit_code"] == expected_exit and result["written_count"] == 0 and result["execution"] is None,
                    "Destination failure announced successful output")
            actual_code, message, diagnostic = result["code"], result["message"], result["diagnostic"]
        require(actual_code == code, "Destination failure category differs: " + identifier + ": " + actual_code)
        preserved = {neighbor.name}
        if target == too_long:
            preserved.update(item.name for item in target.iterdir() if item.name.endswith("-neighbor.txt"))
        if diagnostic:
            require(diagnostic["cleanup_complete"] is True and diagnostic["retained_staging"] is None
                    and diagnostic["cleanup_error"] is None, "Destination failure left an ordinary owned stage")
            path = Path(diagnostic["record_path"])
            require(path.parent.parent == target and path.name == "failure.json", "Destination diagnostic escaped explicit base")
            owner = json.loads((path.parent / ".WinBookSplit-owner.json").read_text(encoding="utf-8"))
            require(owner == {"schema_version": 1, "kind": "failed", "run_id": diagnostic["run_id"]}, "Destination failure ownership differs")
            launchers.remove_known_directory(path.parent, target, {".WinBookSplit-owner.json", "failure.json"}, ".WinBookSplit-owner.json", owner)
        require({item.name for item in target.iterdir()} == preserved, "Destination rejection created unexplained members")
        require(history.file_digest(source) == source_hash and neighbor.read_bytes() == b"Preserve authored failure neighbor\n",
                "Destination rejection modified source or neighbor")
        failures.append({"id": identifier, "passed": True, "method": method, "error_code": actual_code,
                         "error_message": message, "written_count": 0, "successful_final_count": 0, "outputs": [],
                         "cleanup_complete": True, "source_unchanged": True, "neighbor_unchanged": True,
                         "native_lock_verified": lock_verified, "diagnostic": diagnostic,
                         "injected_errno": errno.ENOSPC if identifier == "full-target" else None,
                         "completed_slices_before_failure": len(calls) - 1 if identifier == "full-target" else None})
    return record, failures


def characterize(work, shells):
    require(os.name == "nt" and len(shells) == 2, "Actual Windows and both explicit supported shell hosts required")
    manual.trusted_original_sources()
    before = runner.source_manifest()
    engine = load("wbs_current_paths_engine", ENGINE)
    generator = history.load_generator()
    fixtures = generator.generate_fixtures(work / "fixtures")
    inputs = {str(path): history.file_digest(path) for path in fixtures.values()}
    hosts, literal, rejected = actual_launchers(work, shells, fixtures, generator, engine, inputs)
    filename = filenames(engine, work, generator, inputs)
    chapter = chapters(engine, work, generator, inputs)
    long_case, failures = destinations(engine, work, generator, inputs)
    observation = history.probe_import(work)
    require(inputs == {path: history.file_digest(Path(path)) for path in inputs}, "Path acceptance changed an immutable authored input")
    require(before == runner.source_manifest(), "Source changed during path acceptance")
    return {"schema_version": 1, "task_id": "M2-T02", "result": "PATH_REGRESSION_PASSED", "success": True, "exit_code": 0,
            "observed_at": datetime.now(timezone.utc).isoformat(), "acceptance_ids": runner.PATH_ACCEPTANCE_IDS,
            "host_cases": hosts, "literal_cases": literal, "rejection_cases": rejected, "filename_case": filename,
            "chapter_cases": chapter, "long_destination_case": long_case, "destination_failure_cases": failures,
            "source_unchanged": True, "baseline_guards_preserved": True, "input_and_neighbor_unchanged": True,
            "machine_settings_unchanged": True, "immutable_original_commit": history.ORIGINAL_COMMIT,
            "tested_path_sha256": before, "input_sha256": inputs, "import_observation": observation,
            "environment": {"python": sys.version, "python_executable": sys.executable, "platform": platform.platform(), "pypdf": history.pypdf.__version__},
            "not_run": ["Explorer drag/drop", "UNC", "arbitrary long-path support", "ACL policy changes", "Calibre conversion", "release package"]}


def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--shell-path", action="append", required=True, type=Path)
    args = parser.parse_args()
    try:
        require(sys.flags.isolated and sys.dont_write_bytecode, "Use the explicit developer Python -I -B")
        path = runner.new_external_path(args.report)
        require(len(args.shell_path) == 2 and all(item.is_absolute() and item.is_file() for item in args.shell_path),
                "Two actual existing absolute supported shell executables required")
        shells = [item.resolve() for item in args.shell_path]
        require(len(set(shells)) == 2, "Distinct actual shell executables required")
        with tempfile.TemporaryDirectory(prefix="WBS-M2-T02-") as directory:
            report = characterize(Path(directory).resolve(), shells)
        report["owned_temp_removed"] = not Path(directory).exists()
        runner.validate_paths_report(report, [str(shell) for shell in shells])
        with path.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(report, stream, indent=2, ensure_ascii=True)
            stream.write("\n")
        print("Path acceptance passed: three literal launchers, eight rejections, nine titles, 120 sections per mode and six destination checks")
        return 0
    except (OSError, RuntimeError, ValueError, subprocess.TimeoutExpired) as error:
        print("Path acceptance failed: " + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
