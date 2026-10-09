"""M2-T01 isolated publication and owned Windows cleanup, authored PDFs only."""

import argparse
from concurrent.futures import ThreadPoolExecutor
from contextlib import redirect_stdout
from datetime import datetime, timezone
from io import StringIO
import importlib.util
import json
import os
from pathlib import Path
import platform
import subprocess
import sys
import tempfile
import time
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("wbs_output_plans", ROOT / "tests/plans/characterize_plan.py")
plans = importlib.util.module_from_spec(spec)
spec.loader.exec_module(plans)
manual, history, require = plans.manual, plans.history, plans.require
runner = manual.load_module("wbs_output_runner", ROOT / "tests/run_tests.py")
ENGINE = ROOT / "engine/winbooksplit_engine.py"
FIXED_TIMESTAMP = "20261009-120000"


def snapshot(directories):
    result = {}
    for directory in directories:
        require(not history.is_reparse(directory), "Prior run became a reparse point")
        for path in directory.iterdir():
            require(path.is_file() and not history.is_reparse(path), "Prior run member changed type")
            result[str(path)] = history.file_digest(path)
    return result


def successful_case(identifier, source, base, process, generator):
    result = manual.result_record(process)
    require(result["status"] == "success" and process["exit_code"] == 0 and not process["stderr"],
            "Actual output CLI failed: " + repr(process))
    execution = result["execution"]
    outputs = manual.published_outputs(base, generator, process)
    pages = len(PdfReader(source).pages)
    manual.check_outputs(outputs, [[0, 3], [3, 6], [6, pages]], pages)
    entries = [{key: value for key, value in item.items() if key not in {"page_count", "sha256", "size_bytes"}}
               for item in execution["outputs"]]
    content = plans.page_content(source)
    actual = [digest for item in outputs for digest in plans.page_content(Path(execution["final_directory"]) / item["filename"])]
    require(actual == content, "Published pages changed synthetic content")
    record = {"id": identifier, "passed": True, "mode": "manual", "total_pages": pages, "outputs": outputs,
              "preview_entries": entries, "coverage": execution["coverage"], "source_identity": execution["source_identity"],
              "written_count": execution["written_count"], "writer_result": execution,
              "original_page_content_sha256": content, "page_content_sha256": actual,
              "neighbor_unchanged": True, "no_replanning_or_reopening": True,
              "manifest_validated": True, "prior_outputs_unchanged": True, **process}
    runner.validate_plan_parity(record)
    runner.validate_plan_page_content(record)
    return record


def repeat_cases(work, fixtures, generator):
    base = work / "repeat-base"
    base.mkdir()
    neighbor = base / "synthetic-neighbor.txt"
    neighbor.write_bytes(b"Preserve prior runs and this authored neighbor\n")
    sources = []
    for name, fixture in (("one", "simple10"), ("two", "nested12")):
        directory = work / name
        directory.mkdir()
        source = directory / "Book.pdf"
        source.write_bytes(fixtures[fixture].read_bytes())
        sources.append(source)
    records, finals = [], []
    definitions = [("repeat-first", sources[0]), ("repeat-second", sources[0]),
                   ("same-basename-first", sources[0]), ("same-basename-second", sources[1])]
    for identifier, source in definitions:
        prior = snapshot(finals)
        source_digest = history.file_digest(source)
        process = manual.observe(ENGINE, source, base, work, "manual", "4,7")
        record = successful_case(identifier, source, base, process, generator)
        require(prior == snapshot(finals) and history.file_digest(source) == source_digest
                and neighbor.read_bytes() == b"Preserve prior runs and this authored neighbor\n", "Repeat changed an earlier run/source/neighbor")
        finals.append(Path(record["writer_result"]["final_directory"]))
        require(len(set(finals)) == len(finals), "Repeat reused a published directory")
        records.append(record)
    require({path.name for path in base.iterdir()} == {neighbor.name, *(path.name for path in finals)}, "Repeat left unexpected members")
    return records


def concurrent_cases(work, source, generator):
    base = work / "concurrent-base"
    base.mkdir()
    neighbor = base / "synthetic-neighbor.txt"
    neighbor.write_bytes(b"Concurrent neighbor must remain unchanged\n")
    worker = work / "fixed-timestamp-worker.py"
    worker.write_text("import importlib.util,pathlib,sys,time\n"
        "engine,source,base,ready,gate=sys.argv[1:]\n"
        "spec=importlib.util.spec_from_file_location('wbs_concurrent_engine',engine)\n"
        "module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)\n"
        "module._timestamp=lambda:'20261009-120000'\n"
        "pathlib.Path(ready).write_bytes(b'ready');deadline=time.monotonic()+30\n"
        "while not pathlib.Path(gate).is_file():\n"
        " if time.monotonic()>deadline:raise RuntimeError('Authored launch barrier timed out')\n"
        " time.sleep(0.01)\n"
        "sys.argv=[engine,source,base,'manual','4,7'];sys.exit(module.main())\n", encoding="utf-8")
    gate = work / "concurrent-go"
    commands, processes, ready = [], [], []
    started = datetime.now(timezone.utc).isoformat()
    try:
        for index in range(4):
            ready_path = work / f"concurrent-ready-{index}"
            ready.append(ready_path)
            command = [sys.executable, "-I", "-B", str(worker), str(ENGINE), str(source), str(base), str(ready_path), str(gate)]
            commands.append(command)
            processes.append(subprocess.Popen(command, cwd=work, env=history.clean_environment(work), stdout=subprocess.PIPE, stderr=subprocess.PIPE))
        deadline = time.monotonic() + 30
        while not all(path.is_file() for path in ready):
            require(time.monotonic() < deadline and all(process.poll() is None for process in processes), "Concurrent readiness failed")
            time.sleep(0.01)
        gate.write_bytes(b"go")
        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = [executor.submit(process.communicate, timeout=60) for process in processes]
            captured = [future.result() for future in futures]
        records = []
        for index, (process, (stdout, stderr)) in enumerate(zip(processes, captured)):
            observation = {"command": commands[index], "cwd": str(work), "exit_code": process.returncode,
                           "stdout": stdout.decode("utf-8"), "stderr": stderr.decode("utf-8"), "pid": process.pid,
                           "started_at": started, "completed_at": datetime.now(timezone.utc).isoformat(),
                           "fixed_timestamp": FIXED_TIMESTAMP, "barrier_synchronized": True}
            record = successful_case(f"concurrent-{index}", source, base, observation, generator)
            require("_" + FIXED_TIMESTAMP + "_" in Path(record["writer_result"]["final_directory"]).name, "Concurrent timestamp seam did not run")
            records.append(record)
        finals = [Path(record["writer_result"]["final_directory"]) for record in records]
        require(len(set(finals)) == 4 and len({record["writer_result"]["run_id"] for record in records}) == 4
                and len({record["pid"] for record in records}) == 4, "Concurrent runs reused a reservation/process")
        require({path.name for path in base.iterdir()} == {neighbor.name, *(path.name for path in finals)}
                and neighbor.read_bytes() == b"Concurrent neighbor must remain unchanged\n", "Concurrent run crossed ownership")
        return records
    finally:
        for process in processes:
            if process.poll() is None:
                process.kill()
            process.communicate(timeout=10)


def failure_cases(engine, work, source):
    records = []
    writer, reader = engine.write_slice, engine.PdfReader
    for identifier in sorted(runner.OUTPUT_FAILURE_CODES):
        base = work / identifier
        base.mkdir()
        neighbor = base / "synthetic-neighbor.txt"
        neighbor.write_bytes(b"Failure neighbor remains unchanged\n")
        calls = []

        def write(*args, **kwargs):
            calls.append(str(args[3]))
            if identifier == "mid-write" and len(calls) == 2:
                raise OSError("Authored second-slice write failure")
            writer(*args, **kwargs)
            if identifier == "wrong-page-count":
                short = PdfWriter()
                short.add_page(args[0].pages[0])
                with open(args[3], "wb") as stream:
                    short.write(stream)

        def reopen(path, *args, **kwargs):
            if ".WinBookSplit-stage-" in str(path):
                raise OSError("Authored unreadable staged PDF")
            return reader(path, *args, **kwargs)

        seam = patch.object(engine, "PdfReader", side_effect=reopen) if identifier == "reopen" else \
            patch.object(engine, "_write_manifest", side_effect=OSError("Authored completion manifest failure")) if identifier == "manifest" else \
            patch.object(engine, "_publish_run", side_effect=OSError("Authored atomic publication failure")) if identifier == "promotion" else \
            patch.object(engine, "write_slice", side_effect=write)
        with seam, redirect_stdout(StringIO()):
            result = plans.plain(engine.run_split(source, base, "manual", "4,7"))
        require(result["exit_code"] == 6 and result["written_count"] == 0 and result["execution"] is None
                and result["code"] == runner.OUTPUT_FAILURE_CODES[identifier], "Injected failure lost nonzero category: " + repr(result))
        diagnostic = result["diagnostic"]
        require(diagnostic["cleanup_complete"] is True and diagnostic["retained_staging"] is None
                and diagnostic["cleanup_error"] is None, "Ordinary failure did not remove only its owned stage")
        manual.check_base_members(base, result, [neighbor.name])
        record_path = Path(diagnostic["record_path"])
        require({path.name for path in record_path.parent.iterdir()} == {".WinBookSplit-owner.json", "failure.json"}
                and not any(history.is_reparse(path) for path in record_path.parent.iterdir()), "Failure diagnostic has unexpected members")
        marker = json.loads((record_path.parent / ".WinBookSplit-owner.json").read_text(encoding="utf-8"))
        failed = json.loads(record_path.read_text(encoding="utf-8"))
        require(marker == {"schema_version": 1, "kind": "failed", "run_id": diagnostic["run_id"]}
                and failed["status"] == "failed" and failed["code"] == result["code"]
                and failed["cleanup_complete"] is True and len(failed["message"]) <= 2048,
                "Failure diagnostic ownership/category differs")
        require(neighbor.read_bytes() == b"Failure neighbor remains unchanged\n", "Failure changed neighbor")
        records.append({"id": identifier, "passed": True, "exit_code": 6, "result": result,
                        "successful_final_count": 0, "failure_record": failed, "failure_owner": marker,
                        "cleanup_complete": True, "neighbor_unchanged": True, "source_unchanged": True,
                        "completed_slices_before_failure": 1 if identifier == "mid-write" else None})
    return records


def junction(path, target, work):
    require(path.parent.is_relative_to(work) and target.is_relative_to(work) and not os.path.lexists(path), "Authored junction containment failed")
    shell = Path(os.environ["SystemRoot"]) / "System32/WindowsPowerShell/v1.0/powershell.exe"
    environment = history.clean_environment(work)
    environment.update(WBS_JUNCTION_PATH=str(path), WBS_JUNCTION_TARGET=str(target), PSModulePath=str(shell.parent / "Modules"))
    command = [str(shell), "-NoProfile", "-NonInteractive", "-Command",
               "$ErrorActionPreference='Stop'; New-Item -ItemType Junction -Path $env:WBS_JUNCTION_PATH -Target $env:WBS_JUNCTION_TARGET | Out-Null"]
    result = history.run(command, work, environment=environment)
    require(result["exit_code"] == 0 and history.is_reparse(path) and path.resolve() == target, "Actual owned Windows junction was not created")
    return {**result, "command": command, "reparse_verified": True}


def cleanup_cases(engine, work, source):
    target = work / "owned-junction-target"
    target.mkdir()
    sentinel = target / "outside-sentinel.txt"
    sentinel.write_bytes(b"Never delete the separately owned junction target\n")
    digest = history.file_digest(sentinel)
    records = []
    for identifier in sorted(runner.OUTPUT_CLEANUP_IDS):
        base = work / ("cleanup-" + identifier)
        base.mkdir()
        run, link, foreign = None, None, None
        native = None
        try:
            if identifier == "junction-base":
                link = base / "escape"
                native = junction(link, target, work)
                try:
                    engine.OutputRun(link, source)
                except engine.OutputError as error:
                    code = error.code
                else:
                    raise RuntimeError("Reparse output base was accepted")
                os.rmdir(link)
                link = None
            else:
                run = engine.OutputRun(base, source)
                run.write_owned("owned.txt", b"Known created file only\n")
                code = None
                if identifier == "unexpected-member":
                    foreign = Path(run.stage) / "untracked-neighbor.txt"
                    foreign.write_bytes(b"Untracked synthetic member\n")
                elif identifier == "junction-child":
                    link = Path(run.stage) / "escape"
                    native = junction(link, target, work)
                elif identifier == "manifest-path-not-authority":
                    run.write_owned("WinBookSplit_Manifest.json", json.dumps({"outputs": [{"filename": str(sentinel)}]}).encode())
                elif identifier == "held-marker-tamper":
                    try:
                        with open(Path(run.stage) / ".WinBookSplit-owner.json", "wb") as stream:
                            stream.write(b"tampered")
                    except OSError:
                        code = "native_marker_tamper_prevented"
                    else:
                        raise RuntimeError("Held ownership marker permitted writes")
                elif identifier == "held-stage-replacement":
                    try:
                        Path(run.stage).rename(base / "authored-replacement")
                    except OSError:
                        code = "native_stage_replacement_prevented"
                    else:
                        raise RuntimeError("Held stage permitted replacement")
                if identifier in {"unexpected-member", "junction-child", "mock-reparse"}:
                    seam = patch.object(engine, "_ordinary", side_effect=engine.OutputError("output_ownership_failed", "Authored mocked reparse refusal")) \
                        if identifier == "mock-reparse" else patch.object(engine, "_ordinary", wraps=engine._ordinary)
                    with seam:
                        try:
                            run.cleanup()
                        except engine.OutputError as error:
                            code = error.code
                        else:
                            raise RuntimeError("Unsafe cleanup was accepted")
                    require(Path(run.stage, "owned.txt").read_bytes() == b"Known created file only\n", "Rejected cleanup deleted an owned file")
                    if foreign:
                        require(foreign.read_bytes() == b"Untracked synthetic member\n", "Rejected cleanup deleted a foreign member")
                        foreign.unlink()
                        foreign = None
                    if link:
                        require(history.is_reparse(link) and link.parent == Path(run.stage), "Junction cleanup containment changed")
                        os.rmdir(link)
                        link = None
                require(run.cleanup(), "Known owned stage did not clean after authored recovery")
            require(history.file_digest(sentinel) == digest and {path.name for path in target.iterdir()} == {sentinel.name}, "Cleanup reached outside sentinel")
            require(not list(base.iterdir()), "Cleanup left an owned stage")
            records.append({"id": identifier, "passed": True, "rejected": identifier not in {"ordinary-owned", "manifest-path-not-authority"},
                            "cleaned": True, "owned_stage_removed": True, "outside_sentinel_unchanged": True,
                            "actual_windows": True, "error_code": code, "native_junction": native})
        finally:
            if run:
                run.close()
            # Created junctions are removed by their exact lexical path only.
            if link is not None:
                require(link.parent.is_relative_to(work) and history.is_reparse(link), "Junction final cleanup containment failed")
                os.rmdir(link)
    return records


def characterize(work):
    require(os.name == "nt", "Actual Windows output acceptance required")
    manual.trusted_original_sources()
    before = runner.source_manifest()
    engine = manual.load_module("wbs_current_output_engine", ENGINE)
    generator = history.load_generator()
    fixtures = generator.generate_fixtures(work / "fixtures")
    inputs = {str(path): history.file_digest(path) for path in fixtures.values()}
    repeated = repeat_cases(work, fixtures, generator)
    concurrent = concurrent_cases(work, fixtures["simple10"], generator)
    failures = failure_cases(engine, work, fixtures["simple10"])
    cleanup = cleanup_cases(engine, work, fixtures["simple10"])
    observation = history.probe_import(work)
    guard_path = work / "baseline-must-refuse.json"
    guard = history.run([sys.executable, "-I", "-B", str(ROOT / "tests/baseline/characterize_original.py"), "--report", str(guard_path)],
                        work, environment=history.clean_environment(work))
    require(guard["exit_code"] == 1 and "Original launchers changed" in guard["stderr"] and not guard_path.exists(), "Immutable historical guard weakened")
    require(inputs == {str(path): history.file_digest(path) for path in fixtures.values()}, "Output tests changed original fixture inputs")
    require(before == runner.source_manifest(), "Tested source changed during output acceptance")
    return {"schema_version": 1, "task_id": "M2-T01", "result": "OUTPUT_TRANSACTION_REGRESSION_PASSED", "success": True, "exit_code": 0,
            "observed_at": datetime.now(timezone.utc).isoformat(), "acceptance_ids": runner.OUTPUT_ACCEPTANCE_IDS,
            "repeat_cases": repeated, "concurrent_cases": concurrent, "failure_cases": failures, "cleanup_cases": cleanup,
            "source_unchanged": True, "baseline_guards_preserved": True, "input_and_neighbor_unchanged": True,
            "immutable_original_commit": history.ORIGINAL_COMMIT, "historical_original_guard": guard,
            "tested_path_sha256": before, "input_sha256": inputs, "import_observation": observation,
            "environment": {"python": sys.version, "python_executable": sys.executable, "platform": platform.platform(), "pypdf": history.pypdf.__version__},
            "not_run": ["power-loss recovery", "Calibre conversion", "Explorer", "release package"]}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", required=True, type=Path)
    args = parser.parse_args()
    try:
        require(sys.flags.isolated and sys.dont_write_bytecode, "Use explicit supported developer Python -I -B")
        report_path = runner.new_external_path(args.report)
        with tempfile.TemporaryDirectory(prefix="WinBookSplit-M2-T01-tests-") as directory:
            report = characterize(Path(directory).resolve())
        report["owned_temp_removed"] = not Path(directory).exists()
        runner.validate_output_report(report)
        with report_path.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(report, stream, indent=2, ensure_ascii=True)
            stream.write("\n")
        print("Output acceptance passed: four repeat/same-basename runs, four simultaneous fixed-timestamp processes, five failures and eight ownership checks")
        return 0
    except (OSError, RuntimeError, ValueError, subprocess.TimeoutExpired) as error:
        print("Output acceptance failed: " + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
