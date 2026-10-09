"""Small local test/evidence entry point; never run it against private documents.

Invoke an explicit developer Python with -I -B and an absolute script path.
Reports must be new files outside the checkout. Baseline means reproduction of
known original defects, not corrected-engine or Windows workflow acceptance.
"""

from __future__ import annotations

import argparse
import base64
from datetime import datetime, timezone
import hashlib
from importlib.metadata import PackageNotFoundError, version
import json
import os
from pathlib import Path
import platform
import shutil
import subprocess
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = Path(__file__).resolve()
ACCEPTANCE_IDS = ["AC-009", "AC-010"]
MANUAL_ACCEPTANCE_IDS = [f"AC-{number:03}" for number in range(13, 19)]
MANUAL_ORACLE_IDS = {f"MAN-{number:02}" for number in range(1, 23)}
MANUAL_EXTRA_IDS = {"comma", "zero", "missing", "fullwidth", "embedded-space", "late-invalid",
                    "huge", "long-leading-zeros"}
HISTORICAL_BOOKMARK_IDS = {f"BM-{number:02}" for number in range(1, 4)}
MANUAL_ENTRYPOINT_IDS = {"PS51-unrelated", "PS7-unrelated", "BAT-unrelated",
                         "parallel-PS51-unrelated", "parallel-PS7-unrelated", "parallel-BAT-unrelated"}
GIT_SELECTORS = ("GIT_DIR", "GIT_WORK_TREE", "GIT_INDEX_FILE", "GIT_OBJECT_DIRECTORY",
                 "GIT_ALTERNATE_OBJECT_DIRECTORIES")


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def source_manifest() -> dict[str, str]:
    """Hash actual checked-out bytes; avoid circular evidence/status hashes."""
    paths = [ROOT / name for name in (
        "WinBookSplit.bat", "WinBookSplit.ps1", "README.md", "LICENSE", ".gitignore",
        "requirements.txt", "requirements-dev.txt",
        "tests/README.md", "tests/fixtures/README.md", "tests/baseline/README.md",
        "tests/extraction/README.md", "tests/manual/README.md",
        "docs/codex-v1.0.0/PLAN_ORACLES.json",
        "docs/codex-v1.0.0/ACCEPTANCE_CASES.json",
    )]
    for directory in (ROOT / "tests", ROOT / "engine", ROOT / "tools/codex-handoff"):
        paths.extend(path for path in directory.rglob("*")
                     if path.is_file() and path.suffix in {".py", ".ps1", ".psm1", ".psd1", ".json"}
                     and "__pycache__" not in path.parts)
    return {path.relative_to(ROOT).as_posix(): sha256(path.read_bytes())
            for path in sorted(set(paths)) if path.is_file()}


def manifest_digest(manifest: dict[str, str]) -> str:
    return sha256(json.dumps(manifest, sort_keys=True, separators=(",", ":")).encode())


def run_command(argv: list[str], cwd: Path, *, environment: dict[str, str] | None = None,
                timeout: float = 300) -> dict:
    """communicate drains both pipes; argv never goes through shell=True."""
    record = {"command": argv, "cwd": str(cwd), "timeout_seconds": timeout}
    try:
        result = subprocess.run(argv, cwd=cwd, env=environment, capture_output=True,
                                timeout=timeout, check=False)
        stdout, stderr = result.stdout, result.stderr
        record.update(exit_code=result.returncode, timed_out=False)
    except subprocess.TimeoutExpired as error:
        stdout, stderr = error.stdout or b"", error.stderr or b""
        record.update(exit_code=124, timed_out=True)
    except OSError as error:
        stdout, stderr = b"", str(error).encode("utf-8", errors="replace")
        record.update(exit_code=125, timed_out=False)
    record.update(stdout=stdout.decode("utf-8", errors="replace"),
                  stderr=stderr.decode("utf-8", errors="replace"),
                  stdout_sha256=sha256(stdout), stderr_sha256=sha256(stderr))
    return record


def new_external_path(path: Path) -> Path:
    if not path.is_absolute():
        raise ValueError("Report must be an absolute path")
    if os.path.lexists(path):
        raise ValueError("Report already exists; never overwrite evidence")
    resolved = path.resolve()
    if resolved.is_relative_to(ROOT):
        raise ValueError("Report must be outside the checkout")
    if not resolved.parent.is_dir():
        raise ValueError("Report parent directory must already exist")
    return resolved


def package_versions() -> dict[str, str]:
    packages = {}
    for name in ("pypdf", "reportlab", "Pillow", "charset-normalizer"):
        try:
            packages[name] = version(name)
        except PackageNotFoundError:
            packages[name] = "NOT_INSTALLED"
    return packages


def child_environment(work: Path | None = None) -> dict[str, str]:
    environment = os.environ.copy()
    # Provenance must name this checkout even if the caller used Git selectors.
    for name in GIT_SELECTORS:
        environment.pop(name, None)
    if work is not None:
        environment.update(TEMP=str(work), TMP=str(work))
    return environment


def python_suite() -> int:
    # Isolated Python excludes the script/CWD from sys.path. unittest gets only
    # this explicit trusted test directory; fixture helpers use absolute loads.
    tests = str(ROOT / "tests/python")
    sys.path.insert(0, tests)
    suite = unittest.defaultTestLoader.discover(tests, pattern="test_*.py",
                                               top_level_dir=tests)
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    return 0 if result.wasSuccessful() and result.testsRun else 1


def validate_manual_entrypoints(child: dict, requested_shells: list[str]) -> None:
    """Requested actual hosts must have complete evidence, not only a total."""
    entrypoints = child.get("entrypoints")
    if not isinstance(entrypoints, dict):
        raise ValueError("Manual report omitted requested actual entrypoint evidence")
    probes = entrypoints.get("probes")
    # WindowsPath comparison accepts SystemRoot's casing and slash differences
    # while still requiring the requested executable locations.
    requested_paths = {Path(shell) for shell in requested_shells}
    if len(requested_shells) != 2 or len(requested_paths) != 2 \
            or type(entrypoints.get("probe_count")) is not int or entrypoints["probe_count"] != 6 \
            or type(entrypoints.get("parallel_launch_count")) is not int or entrypoints["parallel_launch_count"] != 3 \
            or entrypoints.get("shared_temp_engine_sentinel_unchanged") is not True \
            or entrypoints.get("owned_document_outputs_removed") is not True \
            or not isinstance(probes, list) or len(probes) != 6 \
            or any(not isinstance(probe, dict) or not isinstance(probe.get("id"), str)
                   or type(probe.get("exit_code")) is not int or probe["exit_code"] != 0
                   or probe.get("input_unchanged") is not True
                   or probe.get("cwd_engine_untouched") is not True
                   or probe.get("owned_neighbor_unchanged") is not True
                   or not isinstance(probe.get("shell_executable"), str)
                   or Path(probe["shell_executable"]) not in requested_paths
                   or not isinstance(probe.get("outputs"), list) or not probe["outputs"]
                   for probe in probes) \
            or {probe["id"] for probe in probes} != MANUAL_ENTRYPOINT_IDS \
            or {Path(probe["shell_executable"]) for probe in probes} != requested_paths:
        raise ValueError("Manual report requires six successful requested-host entrypoints, "
                         "three parallel launches and preserved inputs/neighbors/owned cleanup")


def attach_child_report(step: dict, path: Path, kind: str) -> None:
    """A successful process without its promised evidence is a failed step."""
    try:
        data = path.read_bytes()
        child = json.loads(data.decode("utf-8-sig"))
        if not isinstance(child, dict) or child.get("schema_version") != 1:
            raise ValueError("Expected a schema_version=1 JSON object report")
        if kind == "shell":
            if type(child.get("exit_code")) is not int or type(child.get("success")) is not bool:
                raise ValueError("Shell report requires exit_code and success")
            if child["exit_code"] == 0 and child["success"]:
                pester = child.get("pester")
                if not isinstance(pester, dict) or pester.get("result") != "Passed" \
                        or type(pester.get("total")) is not int or pester["total"] < 1:
                    raise ValueError("Successful shell report requires nonempty passing Pester results")
        elif kind == "baseline":
            if child.get("result") != "ORIGINAL_BEHAVIOR_REPRODUCED" \
                    or type(child.get("engine_case_count")) is not int \
                    or child["engine_case_count"] < 1:
                raise ValueError("Baseline report requires nonempty original characterization")
        elif kind == "extraction":
            cases = child.get("engine_cases")
            import_observation = child.get("import_observation", {})
            if child.get("result") != "EXTRACTION_EQUIVALENCE_REPRODUCED" \
                    or child.get("success") is not True or child.get("exit_code") != 0 \
                    or child.get("engine_case_count") != 14 \
                    or not isinstance(cases, list) or len(cases) != 14 \
                    or any(not isinstance(case, dict) or case.get("equivalent") is not True for case in cases) \
                    or len({case.get("oracle_id") for case in cases}) != 14 \
                    or child.get("source_unchanged") is not True \
                    or not isinstance(import_observation, dict) or import_observation.get("import_safe") is not True \
                    or child.get("immutable_original_commit") != "0de84f367f9bd5ddfa3f408a9c29505d7a39633f":
                raise ValueError("Extraction report requires fourteen equivalent cases, safe import and preserved source")
        elif kind == "manual":
            cases = child.get("engine_cases")
            bookmarks = child.get("historical_bookmark_cases")
            import_observation = child.get("import_observation")
            properties = child.get("seeded_property_cases")
            acceptance_ids = child.get("acceptance_ids")
            extras = child.get("extra_cli_cases")
            writer_samples = child.get("sampled_writer_cases")
            if type(child.get("schema_version")) is not int \
                    or child.get("task_id") != "M1-T02" \
                    or child.get("result") != "MANUAL_REGRESSION_PASSED" \
                    or child.get("success") is not True \
                    or type(child.get("exit_code")) is not int or child["exit_code"] != 0 \
                    or type(child.get("engine_case_count")) is not int or child["engine_case_count"] != 22 \
                    or not isinstance(cases, list) or len(cases) != 22 \
                    or any(not isinstance(case, dict) or not isinstance(case.get("oracle_id"), str)
                           or case.get("passed") is not True for case in cases) \
                    or {case["oracle_id"] for case in cases} != MANUAL_ORACLE_IDS \
                    or not isinstance(extras, list) or len(extras) != 8 \
                    or any(not isinstance(case, dict) or not isinstance(case.get("id"), str)
                           or case.get("passed") is not True for case in extras) \
                    or {case["id"] for case in extras} != MANUAL_EXTRA_IDS \
                    or not isinstance(writer_samples, list) or len(writer_samples) != 25 \
                    or any(not isinstance(case, dict) or type(case.get("index")) is not int
                           or not isinstance(case.get("outputs"), list) or not case["outputs"]
                           for case in writer_samples) \
                    or {case["index"] for case in writer_samples} != set(range(25)) \
                    or not isinstance(bookmarks, list) or len(bookmarks) != 3 \
                    or any(not isinstance(case, dict) or not isinstance(case.get("oracle_id"), str)
                           or case.get("equivalent") is not True for case in bookmarks) \
                    or {case["oracle_id"] for case in bookmarks} != HISTORICAL_BOOKMARK_IDS \
                    or not isinstance(acceptance_ids, list) or acceptance_ids != MANUAL_ACCEPTANCE_IDS \
                    or child.get("source_unchanged") is not True \
                    or child.get("baseline_guards_preserved") is not True \
                    or child.get("input_and_neighbor_unchanged") is not True \
                    or child.get("owned_temp_removed") is not True \
                    or child.get("immutable_original_commit") != "0de84f367f9bd5ddfa3f408a9c29505d7a39633f" \
                    or not isinstance(import_observation, dict) or import_observation.get("import_safe") is not True \
                    or not isinstance(properties, dict) or type(properties.get("count")) is not int \
                    or properties["count"] != 250 or properties.get("passed") is not True \
                    or type(properties.get("seed")) is not int or properties["seed"] != 20261009:
                raise ValueError("Manual report requires twenty-two targets, eight extra CLI cases, "
                                 "seeded coverage checks, twenty-five writer samples, three historical bookmark "
                                 "comparisons, safe import, preserved guards/source/inputs and owned cleanup")
            if step.get("requested_shell_paths"):
                validate_manual_entrypoints(child, step["requested_shell_paths"])
        else:
            raise ValueError("Unknown child evidence kind")
        step.update(evidence_sha256=sha256(data), evidence=child)
        if child.get("success") is False or child.get("exit_code", 0) != 0:
            if step["exit_code"] == 0:
                step["exit_code"] = 126
                step["evidence_error"] = "Child report claims failure despite process success"
    except (OSError, ValueError) as error:
        step["evidence_error"] = str(error)
        if step["exit_code"] == 0:
            step["exit_code"] = 126


def shell_command(shell: Path, work: Path, report: Path, tool_root: Path,
                  failure_probe: bool) -> tuple[list[str], dict[str, str]]:
    # Static bootstrap reads only our trusted test source. Paths travel through
    # environment variables as data; no document/path text is evaluated as code.
    bootstrap = (
        "$ErrorActionPreference = 'Stop'; "
        "$block = [scriptblock]::Create([IO.File]::ReadAllText("
        "$env:WBS_TEST_SCRIPT, [Text.Encoding]::UTF8)); "
        "& $block -RepositoryRoot $env:WBS_TEST_ROOT "
        "-ToolRoot $env:WBS_TEST_TOOLS -ReportPath $env:WBS_TEST_REPORT "
        "-WorkRoot $env:WBS_TEST_WORK"
    )
    if failure_probe:
        bootstrap += " -FailureProbe"
    encoded = base64.b64encode(bootstrap.encode("utf-16le")).decode("ascii")
    environment = child_environment(work)
    environment.update(WBS_TEST_SCRIPT=str(ROOT / "tests/Invoke-ShellTests.ps1"),
                       WBS_TEST_ROOT=str(ROOT), WBS_TEST_TOOLS=str(tool_root),
                       WBS_TEST_REPORT=str(report), WBS_TEST_WORK=str(work),
                       TEMP=str(work), TMP=str(work),
                       PSModulePath=str(shell.parent / "Modules"))
    # This applies only to the fresh child. Machine/UserPolicy still takes
    # precedence; never modify persistent policy or request Bypass/elevation.
    return ([str(shell), "-NoProfile", "-NonInteractive", "-ExecutionPolicy",
             "RemoteSigned", "-EncodedCommand", encoded],
            environment)


def execute(args: argparse.Namespace) -> dict:
    before = source_manifest()
    git = shutil.which("git")
    if not git:
        raise ValueError("Git is required for source provenance")
    head = run_command([git, "rev-parse", "HEAD"], ROOT, environment=child_environment())
    tree = run_command([git, "rev-parse", "HEAD^{tree}"], ROOT, environment=child_environment())
    if head["exit_code"] or tree["exit_code"]:
        raise ValueError("Cannot record Git source identity")
    report = {
        "schema_version": 1, "task_id": "M0-T04",
        "observed_at": datetime.now(timezone.utc).isoformat(),
        "layer": args.layer, "failure_probe": args.failure_probe,
        "acceptance_ids": ACCEPTANCE_IDS,
        "acceptance": {key: "NOT_ASSERTED: inspect actual probes/repeat evidence"
                       for key in ACCEPTANCE_IDS},
        "source": {"base_commit": head["stdout"].strip(),
                   "base_tree": tree["stdout"].strip(),
                   "tested_path_sha256": before,
                   "tested_paths_digest": manifest_digest(before),
                   "meaning": "Actual worktree bytes; Git identity alone does not imply a clean checkout"},
        "environment": {"python": sys.version, "python_executable": sys.executable,
                        "python_isolated": bool(sys.flags.isolated),
                        "python_no_bytecode": bool(sys.dont_write_bytecode),
                        "platform": platform.platform(), "machine": platform.machine(),
                        "packages": package_versions()},
        "steps": [],
        "not_run": ["corrected-bookmark acceptance", "full application splitting in both shells",
                    "human Explorer drag/drop", "real Calibre EPUB/AZW3 conversion",
                    "release-package checks"],
    }
    steps = report["steps"]
    with tempfile.TemporaryDirectory(prefix="WinBookSplit-tests-") as directory:
        work = Path(directory).resolve()
        report["owned_work_directory"] = str(work)
        sentinel = work / "synthetic-input-sentinel.txt"
        sentinel.write_text("Original harness data only.\n", encoding="utf-8")
        sentinel_hash = sha256(sentinel.read_bytes())
        environment = child_environment(work)
        if args.failure_probe in {"native", "python"}:
            if args.failure_probe == "native":
                native = Path(os.environ.get("SystemRoot", "C:/Windows")) / "System32/cmd.exe"
                command = [str(native), "/d", "/c", "exit", "23"]
            else:
                command = [sys.executable, "-I", "-B", str(ROOT / "tests/failures/fail_python.py")]
            steps.append({"name": f"known-{args.failure_probe}-failure",
                          **run_command(command, work, environment=environment)})
        elif not args.failure_probe and args.layer in {"python", "full"}:
            steps.append({"name": "stdlib-unittest",
                          **run_command([sys.executable, "-I", "-B", str(SCRIPT),
                                         "--_python-child"], work, environment=environment)})
        if args.layer in {"shell", "full"} or args.failure_probe == "pester":
            for index, shell in enumerate(args.shell_path):
                shell_work = work / f"shell-{index}"
                shell_work.mkdir()
                child_report = shell_work / "report.json"
                command, shell_environment = shell_command(
                    shell, shell_work, child_report, args.tool_root,
                    args.failure_probe == "pester")
                step = {"name": f"shell-{index}", "shell_executable": str(shell),
                        "process_policy_requested": "RemoteSigned (fresh child only)",
                        "builtin_module_path": shell_environment["PSModulePath"],
                        **run_command(command, shell_work, environment=shell_environment)}
                attach_child_report(step, child_report, "shell")
                steps.append(step)
        if not args.failure_probe and args.layer == "baseline":
            child_report = work / "known-original.json"
            step = {"name": "known-original-characterization",
                    "meaning": "Known defects reproduced; no repaired-engine acceptance",
                    **run_command([sys.executable, "-I", "-B",
                                   str(ROOT / "tests/baseline/characterize_original.py"),
                                   "--launcher-probes", "--report", str(child_report)],
                                  work, environment=environment)}
            attach_child_report(step, child_report, "baseline")
            steps.append(step)
        if not args.failure_probe and args.layer == "extraction":
            child_report = work / "extraction-equivalence.json"
            command = [sys.executable, "-I", "-B",
                       str(ROOT / "tests/extraction/characterize_extraction.py"),
                       "--report", str(child_report)]
            for shell in args.shell_path:
                command.extend(["--shell-path", str(shell)])
            step = {"name": "explicit-extraction-equivalence",
                    "meaning": "Immutable original/extracted known defects compared; no repaired-engine acceptance",
                    **run_command(command, work, environment=environment)}
            attach_child_report(step, child_report, "extraction")
            steps.append(step)
        if not args.failure_probe and args.layer in {"manual", "full"}:
            child_report = work / "manual-regression.json"
            command = [sys.executable, "-I", "-B",
                       str(ROOT / "tests/manual/characterize_manual.py"),
                       "--report", str(child_report)]
            for shell in args.shell_path:
                command.extend(["--shell-path", str(shell)])
            step = {"name": "manual-regression",
                    "meaning": "Corrected manual target cases; bookmark behavior compared with historical observations",
                    "requested_shell_paths": [str(shell) for shell in args.shell_path],
                    **run_command(command, work, environment=environment)}
            attach_child_report(step, child_report, "manual")
            steps.append(step)
        if args.failure_probe:
            # Deliberately execute success after failure. Aggregate status is
            # computed from every step, never from LASTEXITCODE/final command.
            steps.append({"name": "success-after-known-failure",
                          **run_command([sys.executable, "-I", "-B", "-c",
                                         "print('synthetic success after failure')"],
                                        work, environment=environment)})
        report["synthetic_input_unchanged"] = sha256(sentinel.read_bytes()) == sentinel_hash
    report["owned_work_directory_removed"] = not work.exists()
    after = source_manifest()
    report["source"]["source_unchanged"] = before == after
    report["success"] = bool(steps) and all(step["exit_code"] == 0 for step in steps) \
        and report["synthetic_input_unchanged"] and report["owned_work_directory_removed"] \
        and report["source"]["source_unchanged"]
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--layer", choices=("python", "shell", "baseline", "extraction", "manual", "full"), default="full")
    parser.add_argument("--report", type=Path, help="New absolute JSON file outside checkout")
    parser.add_argument("--tool-root", type=Path, help="Absolute isolated shell module directory")
    parser.add_argument("--shell-path", type=Path, action="append", default=[],
                        help="Absolute actual powershell.exe/pwsh.exe; repeat for both hosts")
    parser.add_argument("--failure-probe", choices=("native", "python", "pester"))
    parser.add_argument("--_python-child", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args._python_child:
        return python_suite()
    try:
        if not sys.flags.isolated or not sys.dont_write_bytecode:
            raise ValueError("Invoke the explicit developer Python with -I -B")
        if args.report is None:
            raise ValueError("--report is required")
        args.report = new_external_path(args.report)
        if args.layer in {"shell", "full"} or args.failure_probe == "pester":
            if args.tool_root is None or not args.tool_root.is_absolute() or not args.tool_root.is_dir():
                raise ValueError("Shell layer requires an existing absolute --tool-root")
            args.tool_root = args.tool_root.resolve()
            if args.tool_root.is_relative_to(ROOT):
                raise ValueError("Shell tools must be outside the checkout")
            if not args.shell_path:
                raise ValueError("Shell layer requires explicit --shell-path values")
            if any(not shell.is_absolute() or not shell.is_file() for shell in args.shell_path):
                raise ValueError("Every shell path must be an existing absolute executable")
            args.shell_path = [shell.resolve() for shell in args.shell_path]
        elif args.layer in {"extraction", "manual"} and args.shell_path:
            if any(not shell.is_absolute() or not shell.is_file() for shell in args.shell_path):
                raise ValueError("Every integration shell path must be an existing absolute executable")
            args.shell_path = [shell.resolve() for shell in args.shell_path]
        report = execute(args)
        with args.report.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(report, stream, indent=2, ensure_ascii=True)
            stream.write("\n")
        print(f"{'PASS' if report['success'] else 'FAIL'}: {len(report['steps'])} steps; "
              f"report {args.report}")
        return 0 if report["success"] else 1
    except (OSError, ValueError, RuntimeError) as error:
        print(f"Test runner failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
