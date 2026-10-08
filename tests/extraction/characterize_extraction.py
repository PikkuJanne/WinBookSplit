"""M1 mechanical extraction evidence using immutable Git originals and synthetic PDFs.

Known defects must remain equivalent here. This is not repaired-engine acceptance.
Optional actual entrypoint probes own GUID directories in the observed Documents
folder, with markers, strict cleanup guards and no private document enumeration.
"""

from __future__ import annotations

import argparse
import ast
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
import hashlib
import importlib.util
from importlib.metadata import version
import json
import os
from pathlib import Path
import re
import shutil
import stat
import subprocess
import sys
import tempfile
import uuid

import pypdf

ROOT = Path(__file__).resolve().parents[2]
ORIGINAL_COMMIT = "0de84f367f9bd5ddfa3f408a9c29505d7a39633f"
ENGINE = ROOT / "engine/winbooksplit_engine.py"
BASELINE_PATHS = ("WinBookSplit.ps1", "WinBookSplit.bat",
                  "tests/baseline/characterize_original.py", "tests/baseline/expected_original.json",
                  "tests/fixtures/generate_pdf_fixtures.py", "docs/codex-v1.0.0/PLAN_ORACLES.json")


def require(condition, message):
    if not condition:
        raise RuntimeError(message)


def digest(data):
    return hashlib.sha256(data).hexdigest()


def file_digest(path):
    return digest(path.read_bytes())


def clean_environment(work=None):
    environment = os.environ.copy()
    for key in ("GIT_DIR", "GIT_WORK_TREE", "GIT_INDEX_FILE", "GIT_OBJECT_DIRECTORY",
                "GIT_ALTERNATE_OBJECT_DIRECTORIES"):
        environment.pop(key, None)
    for key in ("PYTHONPATH", "PYTHONHOME"):
        environment.pop(key, None)
    if work is not None:
        environment.update(TEMP=str(work), TMP=str(work), PYTHONDONTWRITEBYTECODE="1",
                           PYTHONNOUSERSITE="1", PYTHONUTF8="1")
    return environment


def run(argv, cwd, *, environment=None, stdin=None, timeout=60):
    # communicate drains both streams. Every child is bounded and explicitly named.
    result = subprocess.run(argv, cwd=cwd, env=environment, input=stdin,
                            capture_output=True, text=True, encoding="utf-8",
                            errors="replace", timeout=timeout, check=False)
    return {"exit_code": result.returncode, "stdout": result.stdout, "stderr": result.stderr}


class EntryPointFailure(RuntimeError):
    def __init__(self, message, *, cleanup_safe, pid):
        super().__init__(message)
        self.cleanup_safe = cleanup_safe
        self.pid = pid


def run_entrypoint(argv, cwd, *, environment, stdin, timeout=60):
    """Bound actual PS/BAT children; uncertain termination forbids output cleanup."""
    child = subprocess.Popen(argv, cwd=cwd, env=environment, stdin=subprocess.PIPE,
                             stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                             text=True, encoding="utf-8", errors="replace")
    streams_finished = False
    try:
        try:
            stdout, stderr = child.communicate(input=stdin, timeout=timeout)
            streams_finished = True
            return {"exit_code": child.returncode, "stdout": stdout, "stderr": stderr}
        except subprocess.TimeoutExpired:
            if child.poll() is not None:
                raise EntryPointFailure("Entrypoint timeout: parent already exited; descendant state is uncertain",
                                        cleanup_safe=False, pid=child.pid)
            taskkill = Path(os.environ["SystemRoot"]) / "System32/taskkill.exe"
            try:
                killed = subprocess.run([str(taskkill), "/PID", str(child.pid), "/T", "/F"],
                                        capture_output=True, timeout=10, check=False)
                if killed.returncode != 0:
                    raise RuntimeError("taskkill returned " + str(killed.returncode))
                # Both redirected streams must reach EOF, and the tracked
                # parent must terminate, before deleting any owned output.
                child.communicate(timeout=10)
                require(child.returncode is not None, "Tracked parent did not terminate")
                streams_finished = True
            except (OSError, RuntimeError, subprocess.TimeoutExpired) as error:
                raise EntryPointFailure("Entrypoint timeout: tree termination could not be verified: " + str(error),
                                        cleanup_safe=False, pid=child.pid) from error
            raise EntryPointFailure("Entrypoint timed out; tracked PID tree was terminated and drained",
                                    cleanup_safe=True, pid=child.pid)
    except EntryPointFailure:
        raise
    except Exception as error:
        raise EntryPointFailure("Entrypoint supervision failed; process state is uncertain: " + str(error),
                                cleanup_safe=False, pid=child.pid) from error
    finally:
        # Windows reader threads may still own pipe locks in an uncertain tree.
        # Do not block on close or mask the unsafe-cleanup failure classifier;
        # failed-orchestrator exit releases those handles instead.
        if streams_finished:
            try:
                for stream in (child.stdin, child.stdout, child.stderr):
                    if stream is not None:
                        stream.close()
            except OSError as error:
                raise EntryPointFailure("Terminated entrypoint stream close failed: " + str(error),
                                        cleanup_safe=True, pid=child.pid) from error


def original_sources():
    git = shutil.which("git")
    require(git is not None, "Git is required for the immutable reference")
    sources = {}
    for path in BASELINE_PATHS:
        result = subprocess.run([git, "cat-file", "blob", ORIGINAL_COMMIT + ":" + path],
                                cwd=ROOT, env=clean_environment(), capture_output=True,
                                timeout=30, check=False)
        require(result.returncode == 0, "Cannot read trusted original Git blob: " + path)
        sources[path] = result.stdout
    expected = json.loads(sources["tests/baseline/expected_original.json"])
    require({path: digest(sources[path]) for path in expected["source_sha256"]}
            == expected["source_sha256"], "Trusted Git originals do not match historical raw-byte guards")
    for path in BASELINE_PATHS[1:]:
        require((ROOT / path).read_bytes() == sources[path], "M1 must preserve original test/input artifact: " + path)
    return sources


def embedded_body(source):
    text = source.decode("utf-8-sig").replace("\r\n", "\n")
    bodies = re.findall(r'^\$pyScriptContent = @"\n(.*?)\n"@$', text, re.MULTILINE | re.DOTALL)
    require(len(bodies) == 1, "Exactly one original embedded engine required")
    body = bodies[0] + "\n"
    require("$" not in body and "`" not in body, "Original body contains PowerShell interpolation")
    return body


def function_asts(source):
    return {node.name: ast.dump(node, include_attributes=False)
            for node in ast.parse(source).body if isinstance(node, ast.FunctionDef)}


def load_generator():
    path = ROOT / "tests/fixtures/generate_pdf_fixtures.py"
    spec = importlib.util.spec_from_file_location("wbs_extraction_fixtures", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def outputs(directory, generator):
    records = []
    for path in sorted(directory.glob("*.pdf")):
        ids = generator.page_ids(path)
        require(bool(ids), "Empty generated output: " + path.name)
        records.append({"filename": path.name, "page_ids": ids,
                        "range": [ids[0] - 1, ids[-1]], "sha256": file_digest(path)})
    return records


def probe_import(work):
    directory = work / "import-no-side-effects"
    directory.mkdir()
    report_path = directory / "observation.json"
    environment = clean_environment(directory)
    environment.update(WBS_IMPORT_ENGINE=str(ENGINE), WBS_IMPORT_REPORT=str(report_path))
    code = (
        "import contextlib,importlib.util,io,json,logging,os,pathlib,sys,pypdf; "
        "logger=logging.getLogger('pypdf'); logger.setLevel(logging.WARNING); "
        "before=(logger.level,logger.handlers[:],sys.stdout.line_buffering,sys.stdout.write_through); "
        "before_files=sorted(p.name for p in pathlib.Path.cwd().iterdir()); "
        "spec=importlib.util.spec_from_file_location('wbs_import_probe',os.environ['WBS_IMPORT_ENGINE']); "
        "module=importlib.util.module_from_spec(spec); captured=io.StringIO(); "
        "original_stdout=sys.stdout; original_stderr=sys.stderr; "
        "\nwith contextlib.redirect_stdout(captured),contextlib.redirect_stderr(captured):\n"
        " spec.loader.exec_module(module)\n"
        "assert captured.getvalue()==''; assert sys.stdout is original_stdout and sys.stderr is original_stderr; "
        "assert before==(logger.level,logger.handlers[:],sys.stdout.line_buffering,sys.stdout.write_through); "
        "assert before_files==sorted(p.name for p in pathlib.Path.cwd().iterdir()); "
        "assert all(callable(getattr(module,n,None)) for n in ('split_pdf','write_slice','main')); "
        "pathlib.Path(os.environ['WBS_IMPORT_REPORT']).write_text(json.dumps({'import_safe':True,"
        "'stdout_stderr_empty':True,'cwd_unchanged':True,'logger_stdout_configuration_unchanged':True}),encoding='utf-8')"
    )
    result = run([sys.executable, "-I", "-B", "-c", code], directory, environment=environment)
    require(result["exit_code"] == 0 and not result["stdout"] and not result["stderr"],
            "Engine import has side effects or failed: " + repr(result))
    return {**json.loads(report_path.read_text(encoding="utf-8")), **result}


def compare_engines(work, sources):
    original = work / "original-from-immutable-git.py"
    body = embedded_body(sources["WinBookSplit.ps1"])
    original.write_text(body, encoding="utf-8", newline="\n")
    prior_asts, current_asts = function_asts(body), function_asts(ENGINE.read_text(encoding="utf-8"))
    require(set(prior_asts) == {"log", "split_pdf", "write_slice"}, "Unexpected original callable surface")
    require(all(current_asts.get(name) == tree for name, tree in prior_asts.items()),
            "Extracted processing function AST differs from original")
    generator = load_generator()
    fixtures = generator.generate_fixtures(work / "fixtures")
    expected = json.loads(sources["tests/baseline/expected_original.json"])
    oracles = json.loads(sources["docs/codex-v1.0.0/PLAN_ORACLES.json"])
    oracle_map = {case["id"]: case for group in ("manual", "bookmarks") for case in oracles[group]}
    cases = []
    for case in expected["cases"]:
        case_work = work / case["oracle_id"]
        case_work.mkdir()
        input_path = fixtures[case["fixture"]]
        input_before = file_digest(input_path)
        neighbor = case_work / "neighbor.txt"
        neighbor.write_bytes(b"Original synthetic extraction neighbor\n")
        oracle = oracle_map[case["oracle_id"]]
        mode = "manual" if "input" in oracle else str(oracle["level"])
        observations = {}
        for name, engine in (("original", original), ("extracted", ENGINE)):
            output = case_work / name
            output.mkdir()
            argv = [sys.executable, "-I", "-B", str(engine), str(input_path), str(output), mode]
            if mode == "manual":
                argv.append(oracle["input"])
            result = run(argv, case_work, environment=clean_environment(case_work))
            slices = outputs(output, generator)
            require(len(slices) == len(list(output.iterdir())), "Unexpected non-PDF engine output")
            require(result["exit_code"] == case["exit_code"], "Unexpected understood original exit")
            require([item["range"] for item in slices] == case["ranges"], "Original known-bad ranges changed")
            observations[name] = {**result, "outputs": slices}
        require(observations["original"] == observations["extracted"],
                "Original/extracted page identities, filenames, bytes, stdout/stderr or exit differ: " + case["oracle_id"])
        require(input_before == file_digest(input_path) and neighbor.read_bytes()
                == b"Original synthetic extraction neighbor\n", "Input or neighbor changed")
        cases.append({"oracle_id": case["oracle_id"], "known_defect": case["defect"],
                      "equivalent": True, "actual_ranges": case["ranges"], **observations})
    return {"function_asts_identical": list(prior_asts), "embedded_engine_sha256": file_digest(original),
            "engine_cases": cases, "engine_case_count": len(cases),
            "fixtures": {name: {"sha256": file_digest(path), "page_ids": generator.page_ids(path)}
                         for name, path in fixtures.items()}}, fixtures, generator


def owned_document_directory(documents, token):
    directory = documents / ("wbs-m1-" + token + "_Chapters")
    require(directory.parent.resolve() == documents and not os.path.lexists(directory),
            "Synthetic output ownership preflight failed")
    directory.mkdir()
    (directory / ".wbs-m1-owner.json").write_text(json.dumps({"token": token}), encoding="utf-8")
    (directory / ".wbs-m1-neighbor.txt").write_bytes(b"Original synthetic Documents neighbor\n")
    return directory


def is_reparse(path):
    details = path.lstat()
    return bool(getattr(details, "st_file_attributes", 0) & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0x400))


def cleanup_document_directory(directory, documents, token, allowed_pdf_names):
    # Never recurse, follow reparse points, or delete a directory based solely on
    # its name. Validate resolved containment, exact ownership and every member.
    require(directory.resolve(strict=True).parent == documents
            and directory.name == "wbs-m1-" + token + "_Chapters", "Cleanup containment failed")
    require(not is_reparse(directory), "Output directory became a reparse point")
    members = list(directory.iterdir())
    for member in members:
        require(not is_reparse(member) and stat.S_ISREG(member.lstat().st_mode),
                "Cleanup refuses directories/reparse points")
        require(member.name in {".wbs-m1-owner.json", ".wbs-m1-neighbor.txt", *allowed_pdf_names}
                or re.fullmatch(r"WinBookSplit_Log_[0-9]{8}_[0-9]{6}\.txt", member.name),
                "Cleanup refuses unexpected member: " + member.name)
    marker = directory / ".wbs-m1-owner.json"
    require(json.loads(marker.read_text(encoding="utf-8")) == {"token": token}, "Cleanup ownership failed")
    for member in members:
        member.unlink()
    directory.rmdir()
    require(not os.path.lexists(directory), "Owned output directory remained after cleanup")


def launchers(work, fixtures, generator, shells, cases):
    require(os.name == "nt" and len(shells) == 2, "Entry probes need actual PS5.1 and PS7 hosts")
    system_root = Path(os.environ["SystemRoot"])
    ps51 = system_root / "System32/WindowsPowerShell/v1.0/powershell.exe"
    require(ps51.resolve() in [path.resolve() for path in shells], "Actual Windows PowerShell host required")
    require(all(path.is_absolute() and path.is_file() for path in shells), "Explicit valid shell paths required")
    documents_probe = run([str(ps51), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-Command",
                           "[Environment]::GetFolderPath('MyDocuments')"], work,
                          environment={**clean_environment(work), "PSModulePath": str(ps51.parent / "Modules")})
    require(documents_probe["exit_code"] == 0 and documents_probe["stdout"].strip(), "Cannot observe Documents directory")
    documents = Path(documents_probe["stdout"].strip()).resolve(strict=True)
    require(documents.is_dir(), "Observed Documents is not an existing directory")
    shared_temp = work / "shared-temp"
    shared_temp.mkdir()
    old_engine = shared_temp / "WinBookSplit_Engine_v2.py"
    old_engine.write_bytes(b"Original owned shared-temp engine-name sentinel\n")
    original_sentinel = file_digest(old_engine)
    prior_by_id = {case["oracle_id"]: case for case in cases}
    launched = []

    def one(label, host, fixture_name, oracle_id, batch=False):
        token = uuid.uuid4().hex
        expected_output = prior_by_id[oracle_id]["original"]["outputs"]
        allowed_names = {item["filename"] for item in expected_output}
        cwd = work / label
        cwd.mkdir()
        # A counterfeit CWD engine must remain unexecuted and untouched.
        counterfeit = cwd / "engine"
        counterfeit.mkdir()
        decoy = counterfeit / "winbooksplit_engine.py"
        decoy.write_text("raise RuntimeError('untrusted CWD engine executed')\n", encoding="utf-8")
        decoy_hash = file_digest(decoy)
        input_path = cwd / ("wbs-m1-" + token + ".pdf")
        shutil.copyfile(fixtures[fixture_name], input_path)
        input_hash = file_digest(input_path)
        environment = clean_environment(shared_temp)
        environment["PATH"] = os.pathsep.join((str(Path(sys.executable).parent), str(ps51.parent),
                                               str(system_root / "System32"), str(system_root)))
        environment["PSModulePath"] = str(host.parent / "Modules")
        stdin = "M\n4,7\n\n" if oracle_id == "MAN-03" else "2\n\n"
        if batch:
            # Only generated safe-ASCII paths enter the legacy batch command.
            require(not any(c in str(ROOT / "WinBookSplit.bat") + str(input_path)
                            for c in ' %!&|<>^"'), "Batch probe requires safe ASCII paths")
            argv = [str(system_root / "System32/cmd.exe"), "/d", "/c", str(ROOT / "WinBookSplit.bat"), str(input_path)]
        else:
            argv = [str(host), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File",
                    str(ROOT / "WinBookSplit.ps1"), str(input_path)]
        directory = owned_document_directory(documents, token)
        cleanup_safe = True
        try:
            try:
                result = run_entrypoint(argv, cwd, environment=environment, stdin=stdin)
            except EntryPointFailure as error:
                cleanup_safe = error.cleanup_safe
                if not cleanup_safe:
                    raise RuntimeError(str(error) + "; preserve owned marked output " + directory.name
                                       + "; tracked parent PID " + str(error.pid)) from error
                raise
            actual = outputs(directory, generator)
            require(result["exit_code"] == 0 and actual == expected_output,
                    "Actual entrypoint did not reproduce expected slices: " + label + ": " + repr(result))
            require(input_hash == file_digest(input_path) and decoy_hash == file_digest(decoy),
                    "Entrypoint modified synthetic input/CWD neighbor")
            require((directory / ".wbs-m1-neighbor.txt").read_bytes()
                    == b"Original synthetic Documents neighbor\n", "Documents neighbor changed")
            version_result = run([str(host), "-NoProfile", "-Command", "$PSVersionTable.PSVersion.ToString()"],
                                 cwd, environment=environment)
            require(version_result["exit_code"] == 0, "Host version probe failed")
            return {"id": label, "entrypoint": "BAT -> actual Windows PowerShell" if batch else "PowerShell -File",
                    "host": version_result["stdout"].strip(), "shell_executable": str(host),
                    "controlled_python": sys.executable, "builtin_module_path": environment["PSModulePath"],
                    "python_path_first": str(Path(sys.executable).parent),
                    "oracle_id": oracle_id, "outputs": actual, "input_unchanged": True,
                    "cwd_engine_untouched": True, "owned_neighbor_unchanged": True,
                    "output_location": "<observed-Documents>/wbs-m1-<run-GUID>_Chapters",
                    **{key: value.replace(str(documents), "<observed-Documents>")
                       if isinstance(value, str) else value for key, value in result.items()}}
        finally:
            if cleanup_safe:
                cleanup_document_directory(directory, documents, token, allowed_names)

    specifications = [
        ("PS51-unrelated", ps51, "simple10", "MAN-03", False),
        ("PS7-unrelated", next(path for path in shells if path.resolve() != ps51.resolve()), "nested12", "BM-03", False),
        ("BAT-unrelated", ps51, "simple10", "MAN-03", True),
    ]
    for specification in specifications:
        launched.append(one(*specification))
    with ThreadPoolExecutor(max_workers=3) as executor:
        futures = [executor.submit(one, "parallel-" + item[0], *item[1:]) for item in specifications]
        for future in futures:
            launched.append(future.result())
    require(file_digest(old_engine) == original_sentinel, "Shared-temp engine sentinel changed or removed")
    return {"probes": launched, "probe_count": len(launched), "parallel_launch_count": 3,
            "shared_temp_engine_sentinel_unchanged": True, "owned_document_outputs_removed": True,
            "method": "Actual -File/BAT, piped prompts, controlled Python PATH, real observed Documents; GUID-owned marker/sentinel directories only. No prompt/path/host shim."}


def characterize(work, shells):
    sources = original_sources()
    paths = (*BASELINE_PATHS, "engine/winbooksplit_engine.py", "tests/extraction/characterize_extraction.py")
    before = {path: file_digest(ROOT / path) for path in paths}
    comparison, fixtures, generator = compare_engines(work, sources)
    import_observation = probe_import(work)
    guard_report = work / "original-must-refuse.json"
    guard = run([sys.executable, "-I", "-B", str(ROOT / "tests/baseline/characterize_original.py"),
                 "--report", str(guard_report)], work, environment=clean_environment(work))
    require(guard["exit_code"] == 1 and "Original launchers changed" in guard["stderr"]
            and not guard_report.exists(), "Historical original guard was weakened or mislabelled current source")
    entrypoints = launchers(work, fixtures, generator, shells, comparison["engine_cases"]) if shells else None
    require(before == {path: file_digest(ROOT / path) for path in paths}, "Extraction probes changed tested source")
    return {"schema_version": 1, "task_id": "M1-T01", "observed_at": datetime.now(timezone.utc).isoformat(),
            "result": "EXTRACTION_EQUIVALENCE_REPRODUCED", "success": True, "exit_code": 0,
            "acceptance_ids": ["AC-011"] + (["AC-012"] if shells else []),
            "meaning": "Mechanical equivalence of understood known-bad behavior; no bug repair/release acceptance.",
            "immutable_original_commit": ORIGINAL_COMMIT,
            "immutable_original_path_sha256": {path: digest(data) for path, data in sources.items()},
            "tested_path_sha256": before, "source_unchanged": True,
            "environment": {"python": sys.version, "python_executable": sys.executable,
                            "pypdf": pypdf.__version__, "reportlab": version("reportlab")},
            "import_observation": import_observation, "historical_original_guard": guard,
            "entrypoints": entrypoints, **comparison,
            "not_run": ["fixed-engine acceptance", "Explorer drag/drop", "Calibre conversion",
                        "release-package acceptance"] + ([] if shells else ["actual entrypoints"])}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--shell-path", action="append", default=[], type=Path)
    args = parser.parse_args()
    try:
        require(sys.flags.isolated and sys.dont_write_bytecode, "Use the explicit developer Python -I -B")
        require(args.report.is_absolute() and not os.path.lexists(args.report), "New absolute report path required")
        report_path = args.report.resolve()
        require(not report_path.is_relative_to(ROOT) and report_path.parent.is_dir(), "Report must be external")
        with tempfile.TemporaryDirectory(prefix="WinBookSplit-M1-T01-") as directory:
            report = characterize(Path(directory).resolve(), args.shell_path)
        report["owned_temp_removed"] = not Path(directory).exists()
        require(report["owned_temp_removed"], "Owned temp cleanup failed")
        with report_path.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(report, stream, indent=2, ensure_ascii=True)
            stream.write("\n")
        print("Extraction equivalence reproduced: 14 known-bad cases; actual entrypoint probes: "
              + str(report["entrypoints"]["probe_count"] if report["entrypoints"] else 0))
        return 0
    except (OSError, RuntimeError, subprocess.TimeoutExpired) as error:
        print("Extraction characterization failed: " + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
