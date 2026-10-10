"""M3-T02 actual Windows CLI and plan-only PDF/real EPUB/AZW3 acceptance."""

from __future__ import annotations

import argparse
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


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


process_tests = load("wbs_cli_process", ROOT / "tests/process/characterize_process.py")
manual, history, require = process_tests.manual, process_tests.history, process_tests.require
paths, launchers, conversion, runner = process_tests.paths, process_tests.launchers, process_tests.conversion, process_tests.runner
validator = load("wbs_cli_validator", ROOT / "tests/cli/validate_cli_report.py")
COMPLETED_CASES = []
LAST_CASE = None


def digest(path):
    return sha256(path.read_bytes()).hexdigest()


def identity(path):
    details = path.lstat()
    return {"sha256": digest(path), "size_bytes": details.st_size, "device": details.st_dev,
            "inode": details.st_ino, "attributes": details.st_file_attributes}


def run_unavailable_stdin(command, cwd, environment, timeout=60, consent=None):
    """DEVNULL for scripted calls; one NoPause-only control gives explicit consent."""
    started, clock = datetime.now(timezone.utc).isoformat(), time.monotonic()
    child = subprocess.Popen(command, cwd=cwd, env=environment, stdin=subprocess.DEVNULL if consent is None else subprocess.PIPE,
                             stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    try:
        stdout, stderr = child.communicate(input=consent, timeout=timeout)
    except subprocess.TimeoutExpired as error:
        # A failing no-prompt contract is not cleanup authority. Preserve this
        # bounded, authored workspace if the process/tree cannot be proved stopped.
        if child.poll() is None:
            child.terminate()  # Retained native process handle; no PID-name/tree lookup.
        raise history.EntryPointFailure("CLI unexpectedly waited; descendant stop unproved, workspace retained",
                                        cleanup_safe=False, pid=child.pid) from error
    return {"command": command, "cwd": str(cwd), "stdin": "DEVNULL" if consent is None else "PIPE",
            "stdin_utf8": None if consent is None else consent.decode("utf-8"), "started_at": started,
            "finished_at": datetime.now(timezone.utc).isoformat(), "elapsed_seconds": time.monotonic() - clock,
            "timeout_seconds": timeout, "timed_out": False, "pid": child.pid, "exit_code": child.returncode,
            "stdout": stdout.decode("utf-8", errors="strict"), "stderr": stderr.decode("utf-8", errors="strict"),
            "stdout_sha256": sha256(stdout).hexdigest(), "stderr_sha256": sha256(stderr).hexdigest()}


def outcome(stdout):
    rows = [line[len("[OUTCOME] "):] for line in stdout.splitlines() if line.startswith("[OUTCOME] ")]
    require(len(rows) == 1, "Expected one authoritative actual CLI outcome: " + stdout)
    return json.loads(rows[0])


def console_records(stdout, base):
    logs = [line[len("Log: "):].strip() for line in stdout.splitlines() if line.startswith("Log: ")]
    if not logs:
        return None, [], [], None
    require(len(logs) == 1, "Multiple emitted CLI console logs")
    log = Path(logs[0])
    require(log.is_absolute() and log.parent.parent == base and log.name == "console.log", "CLI log escaped explicit output base")
    text = log.read_text(encoding="utf-8")
    rows = [json.loads(line) for line in text.splitlines() if line.startswith('{')]
    require(all(row.get('protocol') == 'winbooksplit.result' for row in rows), 'Unexpected actual CLI log frame')
    processes = [json.loads(line[len("[PROCESS] "):]) for line in text.splitlines() if line.startswith("[PROCESS] ")]
    return log, rows, processes, digest(log)


def expected_pdf(generator, path, *, flat=False, parents_only=False):
    outline = [] if flat else [
        {"title": "Parent Å 日本 [1] & $(literal)", "start_page": 1,
         "children": [] if parents_only else [{"title": "Child One", "start_page": 2, "children": []}]},
        {"title": "Parent Two", "start_page": 3,
         "children": [] if parents_only else [{"title": "Child Two", "start_page": 4, "children": []}]},
        {"title": "Parent Three", "start_page": 5,
         "children": [] if parents_only else [{"title": "Child Three", "start_page": 6, "children": []}]},
    ]
    path.write_bytes(generator._pdf_bytes({"title": "Original safe M3-T02 CLI fixture", "pages": 6, "outline": outline}))


def case_arguments(kind, source, base, missing, calibre):
    common = ["-InputFile", str(source), "-OutputDirectory", str(base), "-PythonPath", sys.executable]
    manual_args = ["-Mode", "Manual", "-StartPages", "1,3,5"]
    auto1 = ["-Mode", "Auto", "-BookmarkLevel", "1"]
    auto2 = ["-Mode", "Auto", "-BookmarkLevel", "2"]
    if kind in {"version", "help"}:
        return ["-Version"] if kind == "version" else []
    selected = manual_args if "manual" in kind else auto2 if "auto2" in kind or kind == "no-plan-parents" else auto1
    if kind in validator.EXECUTIONS | validator.PREVIEWS | validator.NO_PLANS:
        return common + selected + (["-Preview"] if kind in validator.PREVIEWS else []) + (
            ["-NoPause"] if kind == "execute-manual-nopause" else [] if kind == "preview-auto1-without-ni" else ["-NonInteractive"]) + (
            ["-CalibrePath", str(calibre)] if source.suffix != ".pdf" else []) + (
            ["-KeepConvertedPdf"] if kind == "execute-epub-keep" else [])
    variations = {
        "missing-input": ["-OutputDirectory", str(base), "-PythonPath", str(missing), *manual_args, "-NonInteractive"],
        "missing-mode": common + ["-NonInteractive"],
        "missing-auto-level": common + ["-Mode", "Auto", "-NonInteractive"],
        "missing-manual-starts": common + ["-Mode", "Manual", "-NonInteractive"],
        "auto-with-starts": common + auto1 + ["-StartPages", "1", "-NonInteractive"],
        "manual-with-level": common + manual_args + ["-BookmarkLevel", "1", "-NonInteractive"],
        "level-without-mode": common + ["-BookmarkLevel", "1", "-NonInteractive"],
        "starts-without-mode": common + ["-StartPages", "1", "-NonInteractive"],
        "bad-mode": common + ["-Mode", "Unknown", "-NonInteractive"],
        "bad-level": common + ["-Mode", "Auto", "-BookmarkLevel", "3", "-NonInteractive"],
        "bad-manual-token": common + ["-Mode", "Manual", "-StartPages", "1,2oops,3", "-NonInteractive"],
        "empty-manual-token": common + ["-Mode", "Manual", "-StartPages", "1,,3", "-NonInteractive"],
        "version-with-processing": common + manual_args + ["-Version", "-NonInteractive"],
        "preview-with-retention": common + auto1 + ["-Preview", "-KeepConvertedPdf", "-NonInteractive"],
        "preview-missing-mode": common + ["-Preview"],
        "explicit-manual-missing-starts": common + ["-Mode", "Manual", "-NoPause"],
        "unknown-argument": common + manual_args + ["-UnknownCliFlag", "-NonInteractive"],
        "missing-python": ["-InputFile", str(source), "-OutputDirectory", str(base), "-PythonPath", str(missing), *manual_args, "-NonInteractive"],
        "missing-calibre": common + auto1 + ["-CalibrePath", str(missing), "-NonInteractive"],
        "manual-out-of-range": common + ["-Mode", "Manual", "-StartPages", "1,999", "-NonInteractive"],
    }
    return variations[kind]


def application_case(work, host, kind, generator, ebook_sources, references, calibre):
    global LAST_CASE
    directory = work / (host["id"] + "-" + kind)
    directory.mkdir()
    app, books, cwd, base = [directory / name for name in ("a", "b", "c", "o")]
    copied = process_tests.copy_application(app)
    for target in (books, cwd, base):
        target.mkdir()
    if kind == "version":
        # Only the PS entry point/version contract exist; no engine/helpers,
        # application venv, Python/Calibre candidates or input are available.
        for name in process_tests.APPLICATION:
            if name not in {"WinBookSplit.ps1", "engine/WinBookSplit.Outcomes.json"}:
                (app / name).unlink()
        copied = {name: copied[name] for name in ("WinBookSplit.ps1", "engine/WinBookSplit.Outcomes.json")}
    fmt = "epub" if kind in {"preview-epub", "execute-epub-keep", "missing-calibre"} else "azw3" if kind == "preview-azw3" else "pdf"
    source = books / ("Book Åä 日本 [1] & (O'Neil) $(literal)." + fmt)
    if fmt == "pdf":
        expected_pdf(generator, source, flat=kind in {"no-plan-auto1", "no-plan-auto2"}, parents_only=kind == "no-plan-parents")
    else:
        shutil.copyfile(ebook_sources[fmt], source)
    neighbor = source.with_suffix(".pdf") if fmt != "pdf" else books / "owned-neighbor.txt"
    neighbor.write_bytes(b"Preserve exact source sibling; synthetic neighbor only\n")
    prior = base / "prior-output.pdf"
    prior.write_bytes(b"Preserve an existing output-base file\n")
    missing = directory / "missing-dependency.exe"
    arguments = case_arguments(kind, source, base, missing, calibre)
    environment = history.clean_environment(cwd)
    system = Path(os.environ["SystemRoot"])
    environment.update(PATH=os.pathsep.join((str(system / "System32"), str(system))),
        PSModulePath=str(Path(host["shell_executable"]).parent / "Modules"), WBS_CLI_HELP_APP=str(app / "WinBookSplit.ps1"))
    if kind == "help":
        command = [host["shell_executable"], "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-File", str(ROOT / "tests/cli/Probe-Help.ps1")]
    else:
        command = [host["shell_executable"], "-NoProfile", *([] if kind == "execute-manual-nopause" else ["-NonInteractive"]),
                   "-ExecutionPolicy", "RemoteSigned", "-File", str(app / "WinBookSplit.ps1"), *arguments]
    with paths.read_only_source(source):
        before = {"source": identity(source), "neighbor": identity(neighbor), "prior": identity(prior)}
        expected = content = None
        if kind in validator.EXECUTIONS | validator.PREVIEWS:
            mode = "manual" if "manual" in kind else "2" if "auto2" in kind else "1"
            engine = load("wbs_cli_expected_" + host["id"] + "_" + kind.replace("-", "_"), ROOT / "engine/winbooksplit_engine.py")
            reference = source if fmt == "pdf" else Path(references[fmt]["path"])
            expected = manual.plain(engine.preview_plan(engine.prepare_split(reference, mode, "1,3,5" if mode == "manual" else None, output_base=base)))
            content = conversion.page_content(PdfReader(reference))
            if fmt == "pdf":
                require([[entry["start"], entry["end"]] for entry in expected["entries"]] == validator.RANGES[mode], "Independent exact PDF CLI plan changed")
        LAST_CASE = {"id": host["id"] + "-" + kind, "command": command, "cwd": str(cwd),
                     "input_path": str(source), "output_base": str(base), "source_observations_before": before}
        observed = run_unavailable_stdin(command, cwd, environment, timeout=90,
                                        consent=b"Y\nN\n" if kind == "execute-manual-nopause" else None)
        LAST_CASE.update(observed)
        record = {"id": host["id"] + "-" + kind, "kind": kind, "host_version": host["host_version"],
            "shell_executable": host["shell_executable"], "actual_process": True, "parameters": arguments,
            "application_sha256": copied, "application_members": sorted(copied), "input_path": str(source),
            "output_base": str(base), "source_observations_before": before, "source_read_only_observed": True,
            "source_observations_after": {"source": identity(source), "neighbor": identity(neighbor), "prior": identity(prior)},
            "output_members_before": [prior.name], "output_members_after": sorted(item.name for item in base.iterdir()),
            "outcome": None, "plan": None, "expected_entries": None, "expected_ranges": None,
            "expected_content_sha256": content, "final_publication": None, "publication_manifest": None,
            "publication_manifest_sha256": None, "output_content_sha256": None, "conversion_reference": references.get(fmt),
            "retained_sha256_verified": False, "retained_page_content_verified": False,
            "temporary_conversion_absent": None, "engine_records": [], "process_summaries": [], "log_sha256": None, **observed}
        log = None
        if kind == "help":
            record["help"] = json.loads(observed["stdout"])
        elif kind != "version":
            final_outcome = outcome(observed["stdout"])
            record["outcome"] = final_outcome
            frame = final_outcome.get("engine_result")
            log, rows, transports, log_hash = console_records(observed["stdout"], base)
            # Preview intentionally uses memory logging; its validated final frame
            # is the authoritative plan receipt, not a fabricated disk log.
            record.update(engine_records=rows, process_summaries=transports, log_sha256=log_hash,
                console_evidence=launchers.authenticate_console_manifest(log, final_outcome,
                    source_path=source, source_observation=before["source"]) if log is not None else None)
            if kind in validator.PREVIEWS:
                require(isinstance(frame, dict) and frame.get("status") == "preview", "CLI preview did not return an explicit engine preview frame")
                plan = frame.get("plan")
                record.update(plan=plan, expected_entries=expected["entries"], expected_ranges=[[entry["start"], entry["end"]] for entry in expected["entries"]])
                if fmt != "pdf":
                    generated = plan["conversion"]["generated_pdf_identity"]
                    record["temporary_conversion_absent"] = not Path(generated["path"]).exists()
            elif kind in validator.EXECUTIONS:
                execution = frame["execution"]
                final = Path(execution["final_directory"])
                ranges_expected = [[entry["start"], entry["end"]] for entry in expected["entries"]]
                if fmt == "epub":
                    checked, members = conversion.check_publication(base, execution, content)
                    outputs, actual_content = checked["outputs"], checked["page_content_sha256"]
                    record.update(retained_sha256_verified=checked["retained_sha256_verified"],
                        retained_page_content_verified=checked["retained_page_content_verified"])
                else:
                    outputs = manual.published_outputs(base, generator, frame)
                    for item in outputs:
                        item.update(page_count=len(PdfReader(final / item["filename"]).pages), size_bytes=(final / item["filename"]).stat().st_size)
                    actual_content = [item for entry in execution["outputs"] for item in conversion.page_content(PdfReader(final / entry["filename"]))]
                    members = {".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", *(entry["filename"] for entry in execution["outputs"])}
                require([item["range"] for item in outputs] == ranges_expected, "Native CLI wrote unexpected physical ranges")
                require(actual_content == content, "Native CLI reordered, duplicated or omitted physical content")
                record.update(final_publication=outputs, publication_manifest=execution["manifest"],
                    publication_manifest_sha256=digest(final / "WinBookSplit_Manifest.json"), output_content_sha256=actual_content,
                    expected_entries=expected["entries"], expected_ranges=ranges_expected)
                launchers.remove_known_directory(final, base, members,
                    ".WinBookSplit-owner.json", {"schema_version": 1, "kind": "run", "run_id": execution["run_id"]})
            if frame is not None:
                conversion.remove_diagnostic(base, frame)
        if log is not None:
            launchers.remove_console_directory(log, base, final_outcome, source_path=source, source_observation=before["source"])
        after = {"source": identity(source), "neighbor": identity(neighbor), "prior": identity(prior)}
        require(before == record["source_observations_after"] == after, "CLI changed source/neighbor/prior identity, attributes or bytes")
        require({item.name for item in base.iterdir()} == {prior.name}, "CLI or known cleanup left unexplained output members")
        require(copied == {name: digest(app / name) for name in copied}, "CLI changed copied application bytes")
        record.update(source_observations_after_cleanup=after, output_members_after_cleanup=[prior.name],
            source_neighbor_prior_unchanged=True, application_unchanged=True, known_output_cleanup_complete=True, passed=True)
    validator.validate_case(record)
    COMPLETED_CASES.append(record)
    LAST_CASE = record
    return record


def characterize(work, shells, calibre):
    require(os.name == "nt" and len(shells) == 2, "Actual Windows and both explicit supported hosts required")
    manual.trusted_original_sources()
    before = runner.source_manifest()
    probe = process_tests.host_probe(work)
    hosts = [paths.host_observation(shell, work, probe) for shell in shells]
    require({host["id"] for host in hosts} == {"PS51", "PS7"}, "Two actual distinct supported hosts required")
    generator = history.load_generator()
    ebooks = work / "ebooks"
    fixture_report = conversion.fixtures.generate(ebooks, calibre)
    sources = {fmt: ebooks / ("original-three-chapters." + fmt) for fmt in ("epub", "azw3")}
    references = {fmt: conversion.reference_pdf(source, work / ("reference-" + fmt + ".pdf"), str(calibre), work) for fmt, source in sources.items()}
    cases = [application_case(work, host, kind, generator, sources, references, calibre) for host in hosts for kind in validator.KINDS]
    for host in hosts:
        host["policies_after"] = paths.host_observation(Path(host["shell_executable"]), work, probe)["stored_policies"]
        require(host["stored_policies"] == host["policies_after"], "Stored execution policy changed")
    require(before == runner.source_manifest(), "Repository sources changed during CLI acceptance")
    return {"schema_version": 1, "task_id": "M3-T02", "result": "CLI_REGRESSION_PASSED", "success": True, "exit_code": 0,
        "observed_at": datetime.now(timezone.utc).isoformat(), "acceptance_ids": validator.ACCEPTANCE,
        "tested_path_sha256": before, "source_unchanged": True, "baseline_guards_preserved": True,
        "host_cases": hosts, "cases": cases, "case_count": len(cases), "machine_settings_unchanged": True,
        "fixture_provenance": fixture_report, "real_conversion_references": references,
        "environment": {"python": sys.version, "python_executable": sys.executable, "pypdf": history.pypdf.__version__},
        "not_run": ["human Explorer/key presses", "release package", "clean OS"],
        "limits": ["Actual host CLI processes with DEVNULL stdin except the NoPause-only explicit-consent PIPE control; copied shipped application, not GUI or release-package acceptance.",
                   "Only authored synthetic PDFs and offline EPUB/genuine Calibre-generated AZW3 are processed."]}


def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--shell-path", required=True, action="append", type=Path)
    parser.add_argument("--calibre-path", required=True, type=Path)
    args = parser.parse_args()
    try:
        require(sys.flags.isolated and sys.dont_write_bytecode, "Use explicit supported developer Python -I -B")
        target = runner.new_external_path(args.report)
        require(len(args.shell_path) == 2 and len(set(args.shell_path)) == 2 and all(path.is_absolute() and path.is_file() for path in args.shell_path), "Both existing absolute supported hosts required")
        require(args.calibre_path.is_absolute() and args.calibre_path.is_file(), "Actual pinned Calibre executable required")
        temporary = tempfile.TemporaryDirectory(prefix="C-")
        work = Path(temporary.name).resolve()
        try:
            report = characterize(work, [path.resolve() for path in args.shell_path], args.calibre_path.resolve())
        except history.EntryPointFailure as error:
            if error.cleanup_safe:
                temporary.cleanup()
            else:
                temporary._finalizer.detach()
                print("Unproved owned process stop; retained CLI workspace: " + str(work), file=sys.stderr)
            raise
        except BaseException:
            temporary._finalizer.detach()
            print("CLI acceptance failed; retained authored workspace: " + str(work), file=sys.stderr)
            raise
        else:
            temporary.cleanup()
        report["owned_temp_removed"] = not work.exists()
        validator.validate_cli_report(report, [str(path.resolve()) for path in args.shell_path], str(args.calibre_path.resolve()))
        with target.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(report, stream, indent=2, ensure_ascii=True)
            stream.write("\n")
        print("CLI acceptance passed: " + str(report["case_count"]) + " actual PS5.1/PS7 controls including real EPUB/AZW3 previews")
        return 0
    except (OSError, RuntimeError, ValueError, KeyError, TypeError, subprocess.TimeoutExpired) as error:
        if 'target' in locals() and not target.exists():
            with target.open("x", encoding="utf-8", newline="\n") as stream:
                json.dump({"schema_version": 1, "task_id": "M3-T02", "result": "CLI_REGRESSION_FAILED", "success": False,
                    "exit_code": 1, "error": str(error), "completed_cases": COMPLETED_CASES, "last_case": LAST_CASE,
                    "cleanup_safe": not work.exists() if 'work' in locals() else True,
                    "owned_temp_removed": not work.exists() if 'work' in locals() else None}, stream, indent=2, ensure_ascii=True)
                stream.write("\n")
        print("CLI acceptance failed: " + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
