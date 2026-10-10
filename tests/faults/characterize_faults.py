"""M4-T03 bounded combined faults, independent invariants and native source binding."""

import argparse
from datetime import datetime, timezone
import importlib.util
import json
import os
from pathlib import Path
import platform
import stat
import sys
import tempfile

ROOT = Path(__file__).resolve().parents[2]
COMPLETED_SECTIONS = {}
LAST_SECTION = None
ACTIVE_SECTION = None


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, ROOT / path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


runner = load("wbs_fault_runner", "tests/run_tests.py")
validator = load("wbs_fault_validator", "tests/faults/validate_faults_report.py")


def require(value, message):
    if not value:
        raise ValueError(message)


def no_reparse_children(directory):
    """Check exact authored temporary children before recursive fixture removal."""
    stack = [directory]
    visited = 0
    while stack:
        parent = stack.pop()
        require(not os.lstat(parent).st_file_attributes & stat.FILE_ATTRIBUTE_REPARSE_POINT,
                "Retain fault workspace containing a reparse point")
        for child in parent.iterdir():
            info = os.lstat(child)
            visited += 1
            require(visited <= 20000 and not info.st_file_attributes & stat.FILE_ATTRIBUTE_REPARSE_POINT,
                    "Retain unexpected/reparse fixture data")
            if stat.S_ISDIR(info.st_mode):
                stack.append(child)
            else:
                require(stat.S_ISREG(info.st_mode), "Retain unexpected non-file fixture data")


def characterize(work, shells):
    global LAST_SECTION, ACTIVE_SECTION
    before = runner.source_manifest()
    definitions = (
        ("isolation", "tests/faults/characterize_isolation.py", "characterize"),
        ("invariants", "tests/faults/invariants.py", "characterize"),
        ("cross_host", "tests/faults/cross_host.py", "characterize_cross_host"),
    )
    COMPLETED_SECTIONS.clear()
    ACTIVE_SECTION = None
    for name, path, function in definitions:
        LAST_SECTION = name
        module = load("wbs_fault_" + name, path)
        ACTIVE_SECTION = module
        arguments = (work / name, shells) if name == "cross_host" else (work / name,)
        COMPLETED_SECTIONS[name] = getattr(module, function)(*arguments)
        require(before == runner.source_manifest(), "Source changed during " + name)
    report = {"schema_version": 1, "task_id": "M4-T03", "observed_at": datetime.now(timezone.utc).isoformat(),
              "result": "FAULT_REGRESSION_PASSED", "success": True, "exit_code": 0,
              "acceptance_ids": validator.ACCEPTANCE, "sections": dict(COMPLETED_SECTIONS),
              "requested_shell_paths": [str(path) for path in shells],
              "tested_path_sha256": before, "source_unchanged": before == runner.source_manifest(),
              "cleanup_safe": True,
              "cross_host_error_scope": "Requires the same-source paths, process and outcomes parent stages",
              "environment": {"python": sys.version, "python_executable": sys.executable,
                              "python_isolated": bool(sys.flags.isolated), "python_no_bytecode": bool(sys.dont_write_bytecode),
                              "platform": platform.platform(), "packages": runner.package_versions()},
              "not_run": ["human/Explorer", "physical disk exhaustion", "clean OS", "CI", "release package"],
              "limits": ["Injected faults are restricted to authored workers or owned copied applications.",
                         "The full cross-host error result requires the unchanged same-source native path/process/outcome suites.",
                         "Finite samples and a deterministic seed do not establish a same-account security sandbox."]}
    validator.validate_faults_report(report, report["requested_shell_paths"], cleanup_complete=False,
                                     expected_source_sha256=before)
    return report


def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--shell-path", required=True, action="append", type=Path)
    args = parser.parse_args()
    temporary = None
    try:
        require(os.name == "nt" and sys.flags.isolated and sys.dont_write_bytecode,
                "Actual Windows and pinned isolated developer Python -I -B required")
        target = runner.new_external_path(args.report)
        require(len(args.shell_path) == 2 and len(set(args.shell_path)) == 2
                and all(path.is_absolute() and path.is_file() for path in args.shell_path),
                "Both actual supported hosts are required")
        shells = [path.resolve() for path in args.shell_path]
        temporary = tempfile.TemporaryDirectory(prefix="F-")
        work = Path(temporary.name).resolve()
        try:
            report = characterize(work, shells)
            no_reparse_children(work)
        except BaseException:
            temporary._finalizer.detach()
            raise
        temporary.cleanup()
        report["owned_temp_removed"] = not work.exists()
        validator.validate_faults_report(report, [str(path) for path in shells])
        with target.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(report, stream, indent=2, ensure_ascii=True)
            stream.write("\n")
        print("PASS: combined run isolation, independent seeded invariants and both-host captured-source controls")
        return 0
    except BaseException as error:
        # A retained inner workspace must have a failed receipt before the
        # parent considers removing its outer temporary ancestor.
        code = 130 if isinstance(error, KeyboardInterrupt) else \
            error.code if isinstance(error, SystemExit) and type(error.code) is int and error.code != 0 else 1
        if temporary is not None:
            temporary._finalizer.detach()
        if "target" in locals() and not target.exists():
            failed = {"schema_version": 1, "task_id": "M4-T03", "result": "FAULT_REGRESSION_FAILED",
                      "success": False, "exit_code": code, "error": str(error), "error_type": type(error).__name__,
                      "last_section": LAST_SECTION,
                      "failed_section_last_case": getattr(ACTIVE_SECTION, "LAST_CASE", None),
                      "failed_section_completed_cases": getattr(ACTIVE_SECTION, "COMPLETED_CASES", None),
                      "completed_sections": dict(COMPLETED_SECTIONS), "cleanup_safe": False,
                      "workspace_retained": str(work) if "work" in locals() else None,
                      "owned_temp_removed": not work.exists() if "work" in locals() else None,
                      "tested_path_sha256_after_failure": runner.source_manifest()}
            with target.open("x", encoding="utf-8", newline="\n") as stream:
                json.dump(failed, stream, indent=2, ensure_ascii=True)
                stream.write("\n")
        print("Fault acceptance failed: " + str(error), file=sys.stderr)
        return code


if __name__ == "__main__":
    raise SystemExit(main())
