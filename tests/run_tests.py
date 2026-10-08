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
        "tests/extraction/README.md",
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
        "not_run": ["corrected-engine acceptance", "full application splitting in both shells",
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
        if not args.failure_probe and args.layer in {"extraction", "full"}:
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
    parser.add_argument("--layer", choices=("python", "shell", "baseline", "extraction", "full"), default="full")
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
        elif args.layer == "extraction" and args.shell_path:
            if any(not shell.is_absolute() or not shell.is_file() for shell in args.shell_path):
                raise ValueError("Every extraction shell path must be an existing absolute executable")
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
