"""Distinct diagnostic categories and actual PS5.1/PS7 handlers, synthetic inputs only."""

from __future__ import annotations

import argparse
from contextlib import redirect_stdout
from copy import deepcopy
from datetime import datetime, timezone
from io import StringIO
import json
import os
from pathlib import Path
import platform
import subprocess
import sys
import tempfile
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter
from pypdf.generic import ArrayObject, NameObject, NullObject
import importlib.util

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("wbs_diagnostic_helpers", ROOT / "tests/plans/characterize_plan.py")
plans = importlib.util.module_from_spec(spec)
spec.loader.exec_module(plans)
manual, history, require = plans.manual, plans.history, plans.require
runner = manual.load_module("wbs_diagnostic_runner", ROOT / "tests/run_tests.py")
ENGINE = ROOT / "engine/winbooksplit_engine.py"


def parse_result(stdout):
    records = [json.loads(line) for line in stdout.splitlines() if line.startswith("{")]
    require(len(records) == 1 and records[0].get("protocol") == "winbooksplit.result"
            and "[NO_BOOKMARKS_FOUND]" not in stdout, "Exactly one explicit diagnostic protocol result required")
    return records[0]


def fixtures(work, generator):
    sources = generator.generate_fixtures(work / "fixtures")
    flat = work / "flat.pdf"
    flat.write_bytes(generator._pdf_bytes({"title": "Synthetic flat diagnostic PDF", "pages": 10, "outline": []}))
    zero = work / "zero.pdf"
    with zero.open("xb") as stream:
        PdfWriter().write(stream)
    invalid = work / "unusable.pdf"
    writer = PdfWriter()
    for page in PdfReader(flat).pages:
        writer.add_page(page)
    node = writer.add_outline_item("Unusable synthetic destination", 0).get_object()
    node.pop("/A", None)
    node[NameObject("/Dest")] = ArrayObject([NullObject(), NameObject("/Fit")])
    with invalid.open("xb") as stream:
        writer.write(stream)
    bad = {}
    for name, data in (("empty-file", b""), ("corrupt-bytes", b"This is authored corrupt PDF data.\n"),
                       ("truncated-pdf", sources["simple10"].read_bytes()[:80])):
        bad[name] = work / (name + ".pdf")
        bad[name].write_bytes(data)
    cycle = plans.level1.cyclic_pdf(sources["simple10"], work / "cycle.pdf")
    unicode_source = plans.level1.fixture(work / "synthetic unicode å 书.pdf", generator, 10,
        [{"title": "章节 å", "start_page": 4, "children": []}, {"title": "次章 é", "start_page": 7, "children": []}])
    return {**sources, "flat": flat, "zero": zero, "unusable": invalid, "cycle": cycle,
            "unicode10": unicode_source, "missing-file": work / "missing-owned-input.pdf", **bad}


def engine_cases(engine, work, sources, generator):
    records = []
    for identifier, (mode, status, _, _, _) in runner.DIAGNOSTIC_EXPECTATIONS.items():
        cwd = work / identifier
        cwd.mkdir()
        output = cwd / "output"
        output.mkdir()
        neighbor = output / "synthetic-neighbor.txt"
        neighbor.write_bytes(b"Preserve this diagnostic neighbor\n")
        if identifier.startswith("flat-"):
            source = sources["flat"]
        elif identifier.startswith("no-usable"):
            source = sources["unusable"]
        elif identifier.startswith("zero-page"):
            source = sources["zero"]
        elif identifier in {"empty-file", "corrupt-bytes", "truncated-pdf", "missing-file"}:
            source = sources[identifier]
        elif identifier == "malformed-outline":
            source = sources["cycle"]
        elif identifier == "success-level1":
            source = sources["unicode10"]
        else:
            source = sources["nested12"] if identifier == "success-level2" else sources["simple10"]
        input_before = history.file_digest(source) if source.exists() else None
        data = "1,invalid" if identifier == "invalid-manual" else "4,7"
        existing = {}
        if identifier == "existing-output":
            preview = engine.preview_plan(engine.prepare_split(source, "manual", data))
            target = output / preview["entries"][0]["filename"]
            target.write_bytes(source.read_bytes())
            existing[target.name] = history.file_digest(target)
        if identifier in {"invalid-plan", "write-failure"}:
            change = patch.object(engine, "prepare_split", side_effect=engine.PlanError("invalid_plan", "Synthetic rejected invalid plan")) \
                if identifier == "invalid-plan" else patch.object(engine, "write_slice", side_effect=OSError("Synthetic write failure before first output"))
            with change, redirect_stdout(StringIO()) as captured:
                diagnostic = plans.plain(engine.run_split(source, output, "manual", data))
            process = {"exit_code": diagnostic["exit_code"], "stdout": captured.getvalue() + json.dumps(diagnostic) + "\n", "stderr": ""}
        elif identifier in {"unexpected-extra-argument", "missing-arguments"}:
            argv = [sys.executable, "-I", "-B", str(ENGINE)]
            if identifier == "unexpected-extra-argument":
                argv += [str(source), str(output), "manual", "4,7", "unexpected extra data"]
            process = history.run(argv, cwd, environment=history.clean_environment(cwd))
            diagnostic = parse_result(process["stdout"])
        else:
            process = manual.observe(ENGINE, source, output, cwd, mode, data if mode == "manual" else None)
            diagnostic = parse_result(process["stdout"])
        outputs = history.outputs(output, generator) if not existing else []
        record = {"id": identifier, "passed": True, "diagnostic": diagnostic, "outputs": outputs,
                  "input_unchanged": input_before == (history.file_digest(source) if source.exists() else None),
                  "neighbor_unchanged": neighbor.read_bytes() == b"Preserve this diagnostic neighbor\n",
                  "no_automatic_fallback": True, **process}
        if existing:
            require(existing == {name: history.file_digest(output / name) for name in existing}, "Existing output changed")
            record["existing_outputs_preserved"] = True
        if status == "success":
            preview = engine.preview_plan(engine.prepare_split(source, mode, data if mode == "manual" else None))
            expected = [[0, 2], [2, 3], [3, 6], [6, 8], [8, 10], [10, 12]] if mode == "2" else [[0, 3], [3, 6], [6, 10]]
            manual.check_outputs(outputs, expected, preview["total_pages"])
            record.update(total_pages=preview["total_pages"], preview_entries=plans.plain(preview["entries"]))
            if identifier == "success-level1":
                names = ["01 - Front matter.pdf", "02 - 章节 å.pdf", "03 - 次章 é.pdf"]
                require([item["filename"] for item in outputs] == names
                        and all("[Writing] " + name in process["stdout"] for name in names),
                        "Actual Unicode bookmark filenames/stdout were changed or undecodable")
                record["unicode_preserved"] = True
        else:
            require(not outputs and ("[Writing]" not in process["stdout"] or identifier == "write-failure"),
                    "Rejected diagnostic case wrote or attempted a slice: " + identifier)
        require({path.name for path in output.iterdir()} == {neighbor.name, *existing, *(item["filename"] for item in outputs)},
                "Diagnostic case created unexpected output")
        records.append(record)
    return records


def protocol_cases(cases):
    no_plan = deepcopy(next(case for case in cases if case["id"] == "flat-level1")["diagnostic"])
    success = deepcopy(next(case for case in cases if case["id"] == "success-manual")["diagnostic"])
    records = []
    variants = {"missing-result": "plain human log only", "malformed-json": "{broken",
                "wrong-protocol": {**no_plan, "protocol": "wrong"}, "wrong-version": {**no_plan, "version": 2},
                "native-exit-mismatch": no_plan, "zero-output-success": {**success, "written_count": 0},
                "multiple-results": json.dumps(no_plan) + "\n" + json.dumps(no_plan),
                "invalid-fallback": {**no_plan, "fallback_modes": ["1", "manual"]},
                "protocol-array": {**no_plan, "protocol": ["winbooksplit.result"]},
                "protocol-casing": {**no_plan, "protocol": "WinBookSplit.Result"},
                "failed-result-with-outputs": {**no_plan, "written_count": 1, "execution": success["execution"]}}
    missing = deepcopy(no_plan)
    del missing["warnings"]
    null_outputs = deepcopy(success)
    null_outputs["execution"]["outputs"] = None
    missing_coverage = deepcopy(success)
    del missing_coverage["execution"]["coverage"]
    missing_outputs = deepcopy(success)
    del missing_outputs["execution"]["outputs"]
    variants.update({"missing-result-field": missing, "invalid-count-type": {**success, "written_count": True},
                     "null-success-outputs": null_outputs, "missing-success-coverage": missing_coverage,
                     "invalid-warning-type": {**no_plan, "warnings": {}}, "missing-success-outputs": missing_outputs})
    success_variants = {"zero-output-success", "invalid-count-type", "null-success-outputs", "missing-success-coverage", "missing-success-outputs"}
    for identifier, result in variants.items():
        mode = "manual" if identifier in success_variants else "1"
        exit_code = 0 if identifier in success_variants else 1 if identifier == "native-exit-mismatch" else 55
        records.append({"id": identifier, "stdout": result if isinstance(result, str) else json.dumps(result), "mode": mode, "exit_code": exit_code})
    invalid = next(case for case in cases if case["id"] == "invalid-mode")
    records.append({"id": "invalid-request-mode", "stdout": invalid["stdout"], "mode": "3", "exit_code": 1})
    return records


def actual_hosts(work, shells, cases):
    records = []
    for index, shell in enumerate(shells):
        cwd = work / f"host-{index}"
        cwd.mkdir()
        output = cwd / "output"
        output.mkdir()
        neighbor = cwd / "neighbor.txt"
        neighbor.write_bytes(b"Preserve the actual-host synthetic neighbor\n")
        flood_engine = cwd / "flood-engine.py"
        flood_engine.write_text("import json,sys\n"
            "message=json.dumps({'arguments':sys.argv[1:],'python_executable':sys.executable})\n"
            "result={'protocol':'winbooksplit.result','version':1,'mode':'manual','status':'invalid_input',"
            "'code':'invalid_start_pages','message':message,'warnings':[],'fallback_modes':[],"
            "'exit_code':1,'written_count':0,'execution':None}\n"
            "sys.stdout.write('o'*200000+'\\n'); sys.stderr.write('e'*200000); print(json.dumps(result)); sys.exit(1)\n", encoding="utf-8")
        payload = {"helper": str(ROOT / "engine/WinBookSplit.Diagnostics.ps1"), "application": str(ROOT / "WinBookSplit.ps1"),
                   "python": sys.executable, "flood_engine": str(flood_engine), "output": str(output), "log": str(cwd / "streams.log"),
                   "neighbor": str(neighbor), "cases": [{"id": case["id"], "stdout": case["stdout"], "exit_code": case["exit_code"],
                                                         "mode": case["diagnostic"]["mode"]} for case in cases],
                   "choices": runner.DIAGNOSTIC_CHOICES, "native_arguments": runner.DIAGNOSTIC_NATIVE_ARGUMENTS,
                   "protocol_cases": protocol_cases(cases)}
        payload_path, report_path = cwd / "payload.json", cwd / "host-report.json"
        payload_path.write_text(json.dumps(payload, ensure_ascii=True), encoding="utf-8")
        environment = history.clean_environment(cwd)
        environment.update(PATH=str(Path(sys.executable).parent) + os.pathsep + environment.get("PATH", ""),
                           PYTHONUTF8="1", PYTHONIOENCODING="utf-8", PSModulePath=str(shell.parent / "Modules"))
        process = history.run([str(shell), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-File",
                               str(ROOT / "tests/diagnostics/Probe-Diagnostics.ps1"), "-PayloadPath", str(payload_path),
                               "-ReportPath", str(report_path)], cwd, environment=environment)
        require(process["exit_code"] == 0 and report_path.is_file(), "Actual diagnostic host failed: " + repr(process))
        records.append({**json.loads(report_path.read_text(encoding="utf-8-sig")), "shell_executable": str(shell),
                        "process_observation": process, "report_sha256": history.file_digest(report_path)})
    return records


def characterize(work, shells):
    history.original_sources()
    paths = (*history.BASELINE_PATHS, "WinBookSplit.ps1", "engine/winbooksplit_engine.py", "engine/WinBookSplit.Diagnostics.ps1",
             "tests/run_tests.py", "tests/python/test_runner.py", "tests/README.md", "tests/diagnostics/README.md",
             "tests/diagnostics/characterize_diagnostics.py", "tests/diagnostics/Probe-Diagnostics.ps1",
             "tests/bookmarks/characterize_level2.py", "tests/plans/characterize_plan.py", "tests/manual/characterize_manual.py")
    before = {path: history.file_digest(ROOT / path) for path in paths}
    engine = manual.load_module("wbs_diagnostic_current_engine", ENGINE)
    generator = history.load_generator()
    sources = fixtures(work, generator)
    inputs = {str(path): history.file_digest(path) for path in sources.values() if path.exists()}
    cases = engine_cases(engine, work, sources, generator)
    hosts = actual_hosts(work, shells, cases)
    import_observation = history.probe_import(work)
    guard_path = work / "baseline-must-refuse.json"
    guard = history.run([sys.executable, "-I", "-B", str(ROOT / "tests/baseline/characterize_original.py"),
                         "--report", str(guard_path)], work, environment=history.clean_environment(work))
    require(guard["exit_code"] == 1 and "Original launchers changed" in guard["stderr"] and not guard_path.exists(), "Historical guard weakened")
    require(inputs == {path: history.file_digest(Path(path)) for path in inputs} and not sources["missing-file"].exists(), "Diagnostic inputs changed")
    require(before == {path: history.file_digest(ROOT / path) for path in paths}, "Diagnostic source changed during testing")
    return {"schema_version": 1, "task_id": "M1-T06", "result": "DIAGNOSTIC_REGRESSION_PASSED", "success": True, "exit_code": 0,
            "observed_at": datetime.now(timezone.utc).isoformat(), "acceptance_ids": runner.DIAGNOSTIC_ACCEPTANCE_IDS,
            "engine_case_count": len(cases), "engine_cases": cases, "shell_cases": hosts,
            "source_unchanged": True, "baseline_guards_preserved": True, "input_and_neighbor_unchanged": True,
            "immutable_original_commit": history.ORIGINAL_COMMIT, "historical_original_guard": guard,
            "tested_path_sha256": before, "input_sha256": inputs, "import_observation": import_observation,
            "environment": {"python": sys.version, "python_executable": sys.executable, "platform": platform.platform(), "pypdf": history.pypdf.__version__},
            "not_run": ["top-level cancel/retry launcher UI", "Explorer", "Calibre", "release package", "complete output transactions"]}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--shell-path", action="append", required=True, type=Path)
    args = parser.parse_args()
    try:
        require(sys.flags.isolated and sys.dont_write_bytecode, "Use explicit supported developer Python -I -B")
        report_path = runner.new_external_path(args.report)
        require(os.name == "nt" and len(args.shell_path) == 2 and len(set(args.shell_path)) == 2
                and all(shell.is_absolute() and shell.is_file() for shell in args.shell_path), "Both actual supported hosts required")
        with tempfile.TemporaryDirectory(prefix="WinBookSplit-M1-T06-tests-") as directory:
            report = characterize(Path(directory).resolve(), args.shell_path)
        report["owned_temp_removed"] = not Path(directory).exists()
        runner.validate_diagnostics_report(report, [str(shell) for shell in args.shell_path])
        with report_path.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(report, stream, indent=2, ensure_ascii=True)
            stream.write("\n")
        print("Diagnostics passed: 23 engine cases, both actual handlers/choices/protocol guards, native quoting and production dual-stream function")
        return 0
    except (OSError, RuntimeError, ValueError, subprocess.TimeoutExpired) as error:
        print("Diagnostic acceptance failed: " + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
