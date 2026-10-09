"""M3-T03 native PS5.1/PS7/BAT literal source selection and exact menus."""

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
import time

from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[2]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


cli = load("wbs_launcher_cli", ROOT / "tests/cli/characterize_cli.py")
process_tests, history, require = cli.process_tests, cli.history, cli.require
manual, paths, launchers, conversion, runner = cli.manual, cli.paths, cli.launchers, cli.conversion, cli.runner
validator = load("wbs_launcher_validator", ROOT / "tests/launcher/validate_launcher_report.py")
digest, identity = cli.digest, cli.identity
COMPLETED_CASES = []
LAST_CASE = None


def run_prompted(command, cwd, environment, answers, timeout=90):
    """Exact authored UTF-8 input; communicate drains both native streams."""
    started, clock = datetime.now(timezone.utc).isoformat(), time.monotonic()
    child = subprocess.Popen(command, cwd=cwd, env=environment, stdin=subprocess.PIPE,
                             stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    try:
        stdout, stderr = child.communicate(input=answers.encode("utf-8"), timeout=timeout)
    except subprocess.TimeoutExpired as error:
        if child.poll() is None:
            child.terminate()  # Retained process handle; no PID/tree cleanup claim.
        raise history.EntryPointFailure("Launcher waited beyond authored input; descendant stop unproved, workspace retained",
                                        cleanup_safe=False, pid=child.pid) from error
    return {"command": command, "cwd": str(cwd), "stdin": "PIPE_UTF8", "stdin_utf8": answers,
        "stdin_sha256": sha256(answers.encode("utf-8")).hexdigest(), "started_at": started,
        "finished_at": datetime.now(timezone.utc).isoformat(), "elapsed_seconds": time.monotonic() - clock,
        "timeout_seconds": timeout, "timed_out": False, "pid": child.pid, "exit_code": child.returncode,
        "stdout": stdout.decode("utf-8", errors="strict"), "stderr": stderr.decode("utf-8", errors="strict"),
        "stdout_sha256": sha256(stdout).hexdigest(), "stderr_sha256": sha256(stderr).hexdigest()}


def controlled_application(app, kind, batch):
    original = process_tests.copy_application(app)
    modifications = []
    if batch:
        if kind in validator.EXTRA:
            # A launched PS must leave this exact owned sentinel and return73.
            # The BAT itself remains byte-identical to the shipped launcher.
            text = validator.SENTINEL_SOURCE
            modifications = ["copied PS launch sentinel only"]
        else:
            text = (app / "WinBookSplit.ps1").read_text(encoding="utf-8-sig")
            for name in ("OutputDirectory", "PythonPath", "CalibrePath"):
                before = "[string]$" + name + ","
                after = "[string]$" + name + " = $env:WBS_LAUNCHER_" + name.upper() + ","
                require(text.count(before) == 1, "Copied dependency/output control seam changed: " + name)
                text = text.replace(before, after)
                modifications.append("copied parameter default " + name)
            before = "[switch]$NoPause,"
            require(text.count(before) == 1, "Copied NoPause control seam changed")
            text = text.replace(before, "[switch]$NoPause = $true,")
            modifications.append("copied NoPause default true")
        (app / "WinBookSplit.ps1").write_text(text, encoding="utf-8-sig")
    actual = {name: digest(app / name) for name in process_tests.APPLICATION}
    require(actual["WinBookSplit.bat"] == original["WinBookSplit.bat"], "Native test changed shipped BAT bytes")
    return original, actual, modifications


def application_case(work, host, kind, generator, ebooks, references, calibre, *, batch=False):
    global LAST_CASE
    label = ("BAT" if batch else host["id"]) + "-" + kind
    directory = work / label
    directory.mkdir()
    app, books, cwd, base = [directory / name for name in ("a", "b", "c", "o")]
    original, actual, modifications = controlled_application(app, kind, batch)
    for target in (books, cwd, base):
        target.mkdir()
    fmt = "epub" if kind.endswith("epub") else "azw3" if kind.endswith("azw3") else "pdf"
    # Literal tokens name authored data, never commands/environment expansion.
    source = books / ("Å日本 [1] & %WBS_LAUNCHER_EXPAND%!WBS_LAUNCHER_EXPAND! $(literal)." + fmt)
    if fmt == "pdf":
        cli.expected_pdf(generator, source, flat=kind in {"fallback-y", "fallback-n", "fallback-c", "fallback-eof"},
                         parents_only=kind in {"fallback-invalid-manual", "fallback-level1", "fallback-invalid-cancel"})
    else:
        shutil.copyfile(ebooks[fmt], source)
    neighbor = source.with_suffix(".pdf") if fmt != "pdf" else books / "owned-neighbor.txt"
    neighbor.write_bytes(b"Authored source neighbor stays unchanged\n")
    prior = base / "prior-output.pdf"
    prior.write_bytes(b"Authored prior output stays unchanged\n")
    second, third = books / "Second owned.pdf", books / "Third owned.pdf"
    for item in (second, third):
        cli.expected_pdf(generator, item)
    extra_before = {str(item): identity(item) for item in (second, third)}
    answers, mode = validator.recipe(kind, source)
    environment = history.clean_environment(cwd)
    system = Path(os.environ["SystemRoot"])
    environment.update(PATH=os.pathsep.join((str(system / "System32"), str(system))),
        PSModulePath=str(Path(host["shell_executable"]).parent / "Modules"), WBS_LAUNCHER_BAT=str(app / "WinBookSplit.bat"),
        WBS_LAUNCHER_OUTPUTDIRECTORY=str(base), WBS_LAUNCHER_PYTHONPATH=sys.executable,
        WBS_LAUNCHER_CALIBREPATH=str(calibre), WBS_LAUNCHER_INPUT1=str(source), WBS_LAUNCHER_INPUT2=str(second),
        WBS_LAUNCHER_INPUT3=str(third), WBS_LAUNCHER_EXPAND="WRONG_EXPANDED_VALUE", WBS_LAUNCHER_STARTED=str(directory / "ps-started.txt"))
    wrapper_text = wrapper_hash = None
    if batch:
        count = 2 if kind == "two-inputs" else 3 if kind in {"three-inputs", "empty-second-third"} else 0 if kind.startswith("source-") else 1
        parameters = [str(item) for item in (source, second, third)[:count]]
        if kind == "empty-second-third":
            parameters[1] = ""
            environment["WBS_LAUNCHER_INPUT2"] = ""
        wrapper_text = '@echo off\r\nsetlocal ' + ('EnableDelayedExpansion' if kind == "source-delayed-expansion" else 'DisableDelayedExpansion') + '\r\n"%WBS_LAUNCHER_BAT%"' + ''.join(' "%WBS_LAUNCHER_INPUT' + str(index) + '%"' for index in range(1, count + 1)) + '\r\n'
        wrapper = directory / "invoke.cmd"
        wrapper.write_bytes(wrapper_text.encode("ascii"))
        wrapper_hash = digest(wrapper)
        command = [str(system / "System32/cmd.exe"), "/d", "/v:on" if kind == "source-delayed-expansion" else "/v:off", "/c", str(wrapper)]
    else:
        parameters = (["-InputFile", str(source)] if not kind.startswith("source-") else []) + [
            "-OutputDirectory", str(base), "-PythonPath", sys.executable, "-CalibrePath", str(calibre), "-NoPause"]
        command = [host["shell_executable"], "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", str(app / "WinBookSplit.ps1"), *parameters]
    with paths.read_only_source(source):
        before = {"source": identity(source), "neighbor": identity(neighbor), "prior": identity(prior)}
        content = expected = None
        if kind not in validator.CANCELLED | validator.EXTRA:
            reference = source if fmt == "pdf" else Path(references[fmt]["path"])
            engine = load("wbs_launcher_expected_" + label.replace("-", "_"), ROOT / "engine/winbooksplit_engine.py")
            expected = manual.plain(engine.preview_plan(engine.prepare_split(reference, mode, "1,3,5" if mode == "manual" else None, output_base=base)))
            content = conversion.page_content(PdfReader(reference))
        LAST_CASE = {"id": label, "command": command, "cwd": str(cwd), "stdin_utf8": answers,
            "input_path": str(source), "output_base": str(base), "source_observations_before": before}
        observed = run_prompted(command, cwd, environment, answers)
        LAST_CASE.update(observed)
        record = {"id": label, "kind": kind, "host_id": "BAT" if batch else host["id"],
            "host_version": host["host_version"], "shell_executable": host["shell_executable"], "actual_process": True,
            "parameters": parameters, "application_sha256": original, "actual_application_sha256": actual,
            "modifications": modifications, "wrapper_text": wrapper_text, "wrapper_sha256": wrapper_hash,
            "input_path": str(source), "output_base": str(base), "source_read_only_observed": True,
            "source_observations_before": before, "source_observations_after": {"source": identity(source), "neighbor": identity(neighbor), "prior": identity(prior)},
            "additional_inputs_before": extra_before, "additional_inputs_after": {str(item): identity(item) for item in (second, third)},
            "output_members_before": [prior.name], "output_members_after": sorted(item.name for item in base.iterdir()),
            "outcome": None, "engine_records": [], "process_summaries": [], "log_sha256": None, "console_operation_outcomes": [],
            "powershell_launch_sentinel_absent": not (directory / "ps-started.txt").exists(),
            "final_publication": None, "publication_manifest": None, "publication_manifest_sha256": None,
            "expected_entries": expected["entries"] if expected else None,
            "expected_ranges": [[entry["start"], entry["end"]] for entry in expected["entries"]] if expected else None,
            "expected_content_sha256": content, "output_content_sha256": None, "temporary_conversion_absent": None,
            "menu_invalid_message_count": observed["stdout"].count("Choose exactly 1, 2, M or C."),
            "fallback_invalid_message_count": observed["stdout"].count("Choose one of the displayed options."), **observed}
        LAST_CASE = record
        log = None
        if kind not in validator.EXTRA:
            final = cli.outcome(observed["stdout"])
            record["outcome"] = final
            log, rows, transports, log_hash = cli.console_records(observed["stdout"], base)
            if log is None and kind in {"menu-cancel", "menu-eof"}:
                # Preflight may close a console log before the first engine call
                # (which normally prints Log:). Authenticate only the one new
                # ordinary marker-scoped console child in this authored base.
                children = [item for item in base.iterdir() if item.name != prior.name]
                require(len(children) == 1 and re.fullmatch(r"\.WinBookSplit-console-[0-9a-f]{32}", children[0].name)
                    and children[0].is_dir() and not children[0].lstat().st_file_attributes & 1024,
                    "No unique ordinary owned console log for pre-engine cancellation")
                child = children[0]
                marker_path = child / ".WinBookSplit-console-owner.json"
                require(marker_path.is_file() and not marker_path.lstat().st_file_attributes & 1024
                    and json.loads(marker_path.read_text(encoding="utf-8")) == {"run_id": child.name.removeprefix(".WinBookSplit-console-"), "kind": "console"},
                    "Pre-engine cancellation console ownership differs")
                require({item.name for item in child.iterdir()} == {marker_path.name, "console.log"}, "Unexpected pre-engine cancellation log member")
                log = child / "console.log"
                require(log.is_file() and not log.lstat().st_file_attributes & 1024, "Cancellation log is not an ordinary file")
                log_text = log.read_text(encoding="utf-8")
                rows = [json.loads(line) for line in log_text.splitlines() if line.startswith('{')]
                transports = [json.loads(line[len("[PROCESS] "):]) for line in log_text.splitlines() if line.startswith("[PROCESS] ")]
                log_hash = digest(log)
                record["log_discovery"] = "exact new owned console marker before engine"
            record.update(engine_records=rows, process_summaries=transports, log_sha256=log_hash)
            frame = final.get("engine_result")
            if isinstance(frame, dict) and frame.get("status") == "success":
                execution = frame["execution"]
                final_dir = Path(execution["final_directory"])
                if fmt != "pdf":
                    checked, members = conversion.check_publication(base, execution, content)
                    outputs, actual_content = checked["outputs"], checked["page_content_sha256"]
                    record["temporary_conversion_absent"] = not Path(execution["conversion"]["generated_pdf_identity"]["path"]).exists()
                else:
                    outputs = manual.published_outputs(base, generator, frame)
                    for item in outputs:
                        item.update(page_count=len(PdfReader(final_dir / item["filename"]).pages), size_bytes=(final_dir / item["filename"]).stat().st_size)
                    actual_content = [page for entry in execution["outputs"] for page in conversion.page_content(PdfReader(final_dir / entry["filename"]))]
                    members = {".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", *(entry["filename"] for entry in execution["outputs"])}
                require([item["range"] for item in outputs] == record["expected_ranges"] and actual_content == content, "Launcher omitted/reordered/duplicated physical pages")
                record.update(final_publication=outputs, publication_manifest=execution["manifest"],
                    publication_manifest_sha256=digest(final_dir / "WinBookSplit_Manifest.json"), output_content_sha256=actual_content)
                launchers.remove_known_directory(final_dir, base, members, ".WinBookSplit-owner.json",
                    {"schema_version": 1, "kind": "run", "run_id": execution["run_id"]})
            for frame in rows:
                conversion.remove_diagnostic(base, frame)
        if log is not None:
            marker = {"run_id": log.parent.name.removeprefix(".WinBookSplit-console-"), "kind": "console"}
            log_text = log.read_text(encoding="utf-8")
            record["console_operation_outcomes"] = [json.loads(line[len("[OPERATION-OUTCOME] "):])
                for line in log_text.splitlines() if line.startswith("[OPERATION-OUTCOME] ")]
            require(record["console_operation_outcomes"] == [record["outcome"]], "Owned console footer differs from final native outcome")
            record["console_log_scope"] = {"path": str(log), "owner": marker,
                "members": sorted((".WinBookSplit-console-owner.json", "console.log")), "sha256": digest(log)}
            launchers.remove_known_directory(log.parent, base, {".WinBookSplit-console-owner.json", "console.log"}, ".WinBookSplit-console-owner.json", marker)
        after = {"source": identity(source), "neighbor": identity(neighbor), "prior": identity(prior)}
        require(before == record["source_observations_after"] == after and extra_before == record["additional_inputs_after"], "Launcher changed source/neighbor/prior/additional identities")
        require({item.name for item in base.iterdir()} == {prior.name}, "Launcher cleanup left unexplained output members")
        require(actual == {name: digest(app / name) for name in actual}, "Launcher changed copied application bytes")
        record.update(source_observations_after_cleanup=after, output_members_after_cleanup=[prior.name],
            source_neighbor_prior_unchanged=True, application_unchanged=True, known_output_cleanup_complete=True, passed=True)
    validator.validate_case(record)
    COMPLETED_CASES.append(record)
    LAST_CASE = record
    return record


def characterize(work, shells, calibre):
    require(os.name == "nt" and len(shells) == 2, "Both actual Windows supported hosts required")
    manual.trusted_original_sources()
    before = runner.source_manifest()
    probe = process_tests.host_probe(work)
    hosts = [paths.host_observation(shell, work, probe) for shell in shells]
    require({host["id"] for host in hosts} == {"PS51", "PS7"}, "Actual PS5.1/PS7 required")
    generator = history.load_generator()
    ebooks = work / "ebooks"
    fixture_report = conversion.fixtures.generate(ebooks, calibre)
    sources = {fmt: ebooks / ("original-three-chapters." + fmt) for fmt in ("epub", "azw3")}
    references = {fmt: conversion.reference_pdf(source, work / ("reference-" + fmt + ".pdf"), str(calibre), work) for fmt, source in sources.items()}
    cases = [application_case(work, host, kind, generator, sources, references, calibre) for host in hosts for kind in validator.COMMON]
    ps51 = next(host for host in hosts if host["id"] == "PS51")
    cases.extend(application_case(work, ps51, kind, generator, sources, references, calibre, batch=True) for kind in validator.COMMON + validator.BAT_ONLY)
    for host in hosts:
        host["policies_after"] = paths.host_observation(Path(host["shell_executable"]), work, probe)["stored_policies"]
        require(host["stored_policies"] == host["policies_after"], "Stored execution policies changed")
    require(before == runner.source_manifest(), "Repository sources changed during launcher acceptance")
    return {"schema_version": 1, "task_id": "M3-T03", "result": "LAUNCHER_REGRESSION_PASSED", "success": True, "exit_code": 0,
        "observed_at": datetime.now(timezone.utc).isoformat(), "acceptance_ids": validator.ACCEPTANCE,
        "evidence_kind": "AUTOMATED_NATIVE_SUPPLEMENT", "human_explorer_tested": False,
        "tested_path_sha256": before, "source_unchanged": True, "baseline_guards_preserved": True,
        "host_cases": hosts, "cases": cases, "case_count": len(cases), "machine_settings_unchanged": True,
        "fixture_provenance": fixture_report, "real_conversion_references": references,
        "environment": {"python": sys.version, "python_executable": sys.executable, "pypdf": history.pypdf.__version__},
        "not_run": ["human Explorer/key presses", "release package", "clean OS"],
        "limits": ["Actual native processes with authored piped choices supplement human Explorer acceptance.",
                   "BAT bytes are exact; successful BAT controls change only copied dependency/output/NoPause defaults.",
                   "BAT multi-input controls replace only copied PS with a launch sentinel to prove it was not called.",
                   "Only authored synthetic PDF and offline EPUB/real Calibre-generated AZW3 inputs are processed."]}


def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--shell-path", required=True, action="append", type=Path)
    parser.add_argument("--calibre-path", required=True, type=Path)
    args = parser.parse_args()
    try:
        require(sys.flags.isolated and sys.dont_write_bytecode, "Use supported developer Python -I -B")
        target = runner.new_external_path(args.report)
        require(len(args.shell_path) == 2 and len(set(args.shell_path)) == 2 and all(path.is_absolute() and path.is_file() for path in args.shell_path), "Both absolute supported hosts required")
        require(args.calibre_path.is_absolute() and args.calibre_path.is_file(), "Pinned actual Calibre required")
        temporary = tempfile.TemporaryDirectory(prefix="L-")
        work = Path(temporary.name).resolve()
        try:
            report = characterize(work, [path.resolve() for path in args.shell_path], args.calibre_path.resolve())
        except BaseException:
            temporary._finalizer.detach()
            print("Unproved launcher acceptance/cleanup; retained authored workspace: " + str(work), file=sys.stderr)
            raise
        else:
            temporary.cleanup()
        report["owned_temp_removed"] = not work.exists()
        validator.validate_launcher_report(report, [str(path.resolve()) for path in args.shell_path], str(args.calibre_path.resolve()))
        with target.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(report, stream, indent=2, ensure_ascii=True)
            stream.write("\n")
        print("Launcher acceptance passed: " + str(report["case_count"]) + " actual PS5.1/PS7/BAT controls")
        return 0
    except (OSError, RuntimeError, ValueError, KeyError, TypeError, subprocess.TimeoutExpired) as error:
        if 'target' in locals() and not target.exists():
            with target.open("x", encoding="utf-8", newline="\n") as stream:
                json.dump({"schema_version": 1, "task_id": "M3-T03", "result": "LAUNCHER_REGRESSION_FAILED", "success": False,
                    "exit_code": 1, "error": str(error), "completed_cases": COMPLETED_CASES, "last_case": LAST_CASE,
                    "cleanup_safe": not work.exists() if 'work' in locals() else True,
                    "retained_workspace": str(work) if 'work' in locals() and work.exists() else None,
                    "owned_temp_removed": not work.exists() if 'work' in locals() else None}, stream, indent=2, ensure_ascii=True)
                stream.write("\n")
        print("Launcher acceptance failed: " + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
