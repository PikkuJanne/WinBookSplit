"""Strict M4-T03 composite child receipt checks; no new application execution."""

import importlib.util
from pathlib import Path
import re
import sys

ROOT = Path(__file__).resolve().parents[2]
SECTIONS = ("isolation", "invariants", "cross_host")
ACCEPTANCE = ["AC-073", "AC-074", "AC-075"]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, ROOT / path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def require(value, message):
    if not value:
        raise ValueError(message)


def validate_faults_report(report, shell_paths, *, cleanup_complete=True, expected_source_sha256=None):
    require(type(report) is dict and type(report.get("schema_version")) is int
            and report["schema_version"] == 1, "Fault receipt schema required")
    require(report.get("task_id") == "M4-T03" and report.get("result") == "FAULT_REGRESSION_PASSED"
            and report.get("success") is True and type(report.get("exit_code")) is int
            and report["exit_code"] == 0, "Successful fault receipt required")
    require(report.get("acceptance_ids") == ACCEPTANCE, "Exact fault acceptance IDs required")
    require(report.get("source_unchanged") is True and report.get("cleanup_safe") is True,
            "Changed source or unproved safe cleanup")
    require(report.get("cross_host_error_scope") == "Requires the same-source paths, process and outcomes parent stages",
            "Source-binding checks alone do not establish the full cross-host error suite")
    require(type(shell_paths) is list and len(shell_paths) == 2 and len(set(shell_paths)) == 2
            and report.get("requested_shell_paths") == shell_paths, "Both exact requested hosts required")
    source = report.get("tested_path_sha256")
    require(type(source) is dict and source and all(type(key) is str and type(value) is str
            and re.fullmatch(r"[a-f0-9]{64}", value) is not None for key, value in source.items()),
            "Raw source map required")
    require({"engine/winbooksplit_engine.py", "engine/winbooksplit_job.py", "tests/run_tests.py",
             "tests/faults/characterize_faults.py", "tests/faults/validate_faults_report.py"} <= set(source),
            "Fault source map omits required runtime or harness bytes")
    if expected_source_sha256 is not None:
        require(source == expected_source_sha256, "Fault child uses different source bytes from its parent")
    sections = report.get("sections")
    require(type(sections) is dict and set(sections) == set(SECTIONS), "Exactly three fault sections required")
    for name in SECTIONS:
        section = sections[name]
        require(type(section) is dict and section.get("tested_path_sha256") == source
                and section.get("source_unchanged") is True, name + ": unbound or changed source")
    isolation = load("wbs_fault_isolation_receipt", "tests/faults/validate_isolation.py")
    invariants = load("wbs_fault_invariant_receipt", "tests/faults/invariants.py")
    cross_host = load("wbs_fault_cross_host_receipt", "tests/faults/validate_cross_host_report.py")
    isolation.validate_isolation_report(sections["isolation"])
    invariants.validate_invariant_report(sections["invariants"], expected_source_sha256=source)
    cross_host.validate_cross_host_report(sections["cross_host"], shell_paths)
    if cleanup_complete:
        require(report.get("owned_temp_removed") is True, "Fault workspace was not removed after verified shutdown")
