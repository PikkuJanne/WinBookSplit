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
import re
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
LEVEL2_ACCEPTANCE_IDS = [f"AC-{number:03}" for number in range(23, 27)]
LEVEL2_ORACLE_IDS = {"BM-03", "BM-04", "BM-06", "BM-07"}
LEVEL2_HIERARCHY_IDS = {"duplicate-parent-subtrees", "invalid-parent-subtrees", "child-order-and-aliases",
                        "invalid-child-destinations", "deep-and-malformed-outlines", "no-usable-level2"}
LEVEL2_REFERENCE_RANGES = [[0, 2], [2, 3], [3, 6], [6, 8], [8, 10], [10, 12]]
LEVEL2_REFERENCE_TITLES = ["Front matter", "A - Opening pages", "A1", "A2", "B - Opening pages", "B1"]
LEVEL1_ACCEPTANCE_IDS = [f"AC-{number:03}" for number in range(19, 23)]
LEVEL1_ORACLE_IDS = {"BM-01", "BM-02", "BM-05", "BM-08"}
LEVEL1_NORMALIZATION_IDS = {"invalid-destinations", "deep-lineage", "cyclic-outline",
                            "malformed-outline", "traversal-limits"}
PLAN_ACCEPTANCE_IDS = ["AC-027", "AC-028", "AC-029"]
PLAN_STRUCTURAL_IDS = {"gap", "overlap", "empty", "negative", "reversed", "overflow",
                       "zero-pages", "noninteger", "valid-whole-document"}
PLAN_PREVIEW_IDS = {"read-only-preview", "isolated-preview-data"}
PLAN_SOURCE_IDS = {"changed-same-path", "repointed-source"}
PLAN_MODES = {"manual", "1", "2"}
OUTPUT_ACCEPTANCE_IDS = ["AC-032", "AC-033", "AC-034", "AC-035"]
OUTPUT_REPEAT_IDS = {"repeat-first", "repeat-second", "same-basename-first", "same-basename-second"}
OUTPUT_CONCURRENT_IDS = {f"concurrent-{number}" for number in range(4)}
OUTPUT_FAILURE_CODES = {"mid-write": "output_write_failed", "reopen": "output_validation_failed",
                        "wrong-page-count": "output_validation_failed", "manifest": "output_manifest_failed",
                        "promotion": "output_publish_failed"}
OUTPUT_CLEANUP_IDS = {"ordinary-owned", "unexpected-member", "held-marker-tamper", "manifest-path-not-authority",
                      "mock-reparse", "junction-base", "junction-child", "held-stage-replacement"}
PATH_ACCEPTANCE_IDS = ["AC-036", "AC-037", "AC-038", "AC-039"]
CONVERSION_ACCEPTANCE_IDS = ["AC-040", "AC-041", "AC-042"]
RUNTIME_ACCEPTANCE_IDS = ["AC-043", "AC-044", "AC-045"]
CONVERSION_REAL_IDS = {host + "-" + fmt + "-" + retention for host in ("PS51", "PS7")
                       for fmt in ("epub", "azw3") for retention in ("default", "keep")}
CONVERSION_INVALID_IDS = {host + "-" + kind for host in ("PS51", "PS7")
                          for kind in ("missing", "empty", "corrupt", "zero-page")}
CONVERSION_CALIBRE_SHA256 = "f46a01c9b8cd392e1190ce2e50a2a5a48b326c8c3eab670bbcd8b298bede501d"
PATH_LITERAL_IDS = {"PS51-literal", "PS7-literal", "BAT-literal"}
PATH_REJECTION_IDS = {host + "-" + case for host in ("PS51", "PS7")
                      for case in ("directory-pdf", "provider-pdf", "corrupt-pdf", "unreadable-held-file")}
PATH_FILENAME_EXPECTATIONS = {
    "forbidden-control": "01 - AB.pdf", "trailing-dots": "02 - Ending.pdf",
    "reserved-title": "03 - CON.pdf", "punctuation-only": "04 - Section 4.pdf",
    "unicode": "05 - 日本語 åäö 😀.pdf", "whitespace": "06 - spaced words.pdf",
    "case-first": "07 - Same.pdf", "case-second": "08 - same.pdf",
    "long-astral": "09 - " + "Å😀" * 16 + "Å.pdf",
}
PATH_DESTINATION_FAILURE_CODES = {"over-budget": "output_path_too_long",
                                  "unbound-long-base": "output_path_too_long",
                                  "changed-bound-base": "output_destination_changed",
                                  "unwritable-held-base": "output_base_invalid",
                                  "full-target": "output_write_failed"}
DIAGNOSTIC_ACCEPTANCE_IDS = ["AC-030", "AC-031"]
DIAGNOSTIC_EXPECTATIONS = {
    "flat-level1": ("1", "no_plan", "no_bookmarks", 55, ["manual"]),
    "flat-level2": ("2", "no_plan", "no_bookmarks", 55, ["manual"]),
    "no-usable-level1": ("1", "no_plan", "no_usable_bookmarks", 55, ["manual"]),
    "no-usable-level2": ("2", "no_plan", "no_usable_bookmarks", 55, ["manual"]),
    "parents-no-level2": ("2", "no_plan", "no_bookmarks_at_level", 55, ["1", "manual"]),
    "zero-page-manual": ("manual", "invalid_input", "invalid_document", 1, []),
    "zero-page-level1": ("1", "invalid_input", "invalid_document", 1, []),
    "zero-page-level2": ("2", "invalid_input", "invalid_document", 1, []),
    "empty-file": ("1", "read_error", "unreadable_document", 1, []),
    "corrupt-bytes": ("1", "read_error", "unreadable_document", 1, []),
    "truncated-pdf": ("1", "read_error", "unreadable_document", 1, []),
    "missing-file": ("1", "read_error", "unreadable_document", 1, []),
    "invalid-mode": ("3", "invalid_input", "invalid_mode", 1, []),
    "invalid-manual": ("manual", "invalid_input", "invalid_start_pages", 1, []),
    "malformed-outline": ("1", "invalid_input", "invalid_outline", 1, []),
    "existing-output": ("manual", "success", "split_complete", 0, []),
    "invalid-plan": ("manual", "error", "invalid_plan", 1, []),
    "write-failure": ("manual", "write_error", "output_write_failed", 1, []),
    "unexpected-extra-argument": ("manual", "invalid_input", "invalid_arguments", 1, []),
    "missing-arguments": ("", "invalid_input", "invalid_arguments", 1, []),
    "success-manual": ("manual", "success", "split_complete", 0, []),
    "success-level1": ("1", "success", "split_complete", 0, []),
    "success-level2": ("2", "success", "split_complete", 0, []),
}
DIAGNOSTIC_PROTOCOL_IDS = {"missing-result", "malformed-json", "wrong-protocol", "wrong-version",
                           "native-exit-mismatch", "zero-output-success", "multiple-results",
                           "invalid-fallback", "failed-result-with-outputs", "invalid-request-mode",
                           "missing-result-field", "invalid-count-type", "null-success-outputs",
                           "missing-success-coverage", "invalid-warning-type", "missing-success-outputs",
                           "protocol-array", "protocol-casing"}
DIAGNOSTIC_CHOICES = ["", "M", "Y", "1", "N", "C", "yes", "maybe", "m", "y", " n "]
DIAGNOSTIC_NATIVE_ARGUMENTS = ["", "quote' [space] å & (paren)", "C:\\literal\\trailing\\",
                               'embedded"quote', 'backslash\\\\"quote', "%value%!literal!$(data)"]
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
        "tests/extraction/README.md", "tests/manual/README.md", "tests/bookmarks/README.md",
        "tests/plans/README.md",
        "tests/diagnostics/README.md",
        "tests/output/README.md",
        "tests/paths/README.md",
        "tests/conversion/README.md",
        "tests/runtime/README.md",
        "tests/process/README.md",
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
    # -I ignores PYTHONIOENCODING; callable Unicode writer tests still need
    # the same explicit stream contract as the production CLI.
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
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


def validate_level2_launcher_reference(reference: object) -> None:
    if not isinstance(reference, dict) or reference.get("oracle_id") != "BM-03" \
            or reference.get("passed") is not True \
            or type(reference.get("exit_code")) is not int or reference["exit_code"] != 0 \
            or reference.get("expected_ranges") != LEVEL2_REFERENCE_RANGES \
            or reference.get("titles") != LEVEL2_REFERENCE_TITLES:
        raise ValueError("Manual report requires the corrected six-section BM-03 launcher reference")
    outputs = reference.get("outputs")
    if not isinstance(outputs, list) or len(outputs) != 6 \
            or any(not isinstance(record, dict) for record in outputs) \
            or [record.get("range") for record in outputs] != LEVEL2_REFERENCE_RANGES \
            or [record.get("page_ids") for record in outputs] != [list(range(start + 1, end + 1))
                                                                 for start, end in LEVEL2_REFERENCE_RANGES] \
            or [record.get("filename") for record in outputs] != [f"{index:02d} - {title}.pdf"
                                                                   for index, title in enumerate(LEVEL2_REFERENCE_TITLES, 1)]:
        raise ValueError("Corrected BM-03 reference requires all six exact ranges, page identities and titles")


def validate_plan_case_ids(cases: object, required: set[str], name: str, *, key: str = "id") -> list[dict]:
    if not isinstance(cases, list) or len(cases) != len(required) \
            or any(not isinstance(case, dict) or not isinstance(case.get(key), str)
                   or case.get("passed") is not True for case in cases) \
            or {case[key] for case in cases} != required:
        raise ValueError(f"Shared-plan report requires every successful {name} case")
    return cases


def validate_plan_parity(case: dict) -> None:
    """Preview and reopened physical output identities must agree completely."""
    pages, entries, outputs = case.get("total_pages"), case.get("preview_entries"), case.get("outputs")
    coverage, identity = case.get("coverage"), case.get("source_identity")
    if type(pages) is not int or pages < 1 or not isinstance(entries, list) or not entries \
            or not isinstance(outputs, list) or len(outputs) != len(entries) \
            or type(case.get("written_count")) is not int or case["written_count"] != len(entries) \
            or not isinstance(coverage, dict) or coverage.get("complete") is not True \
            or type(coverage.get("covered_pages")) is not int or type(coverage.get("section_count")) is not int \
            or coverage != {"complete": True, "covered_pages": pages, "section_count": len(entries)} \
            or case.get("neighbor_unchanged") is not True or case.get("no_replanning_or_reopening") is not True \
            or not isinstance(identity, dict) or identity.get("binding") != "reader_snapshot" \
            or not isinstance(identity.get("path"), str) or not identity["path"] \
            or type(identity.get("size_bytes")) is not int or identity["size_bytes"] < 1 \
            or not isinstance(identity.get("sha256"), str) or len(identity["sha256"]) != 64 \
            or any(character not in "0123456789abcdef" for character in identity["sha256"]):
        raise ValueError("Shared-plan parity requires nonempty complete preview and writer records")
    previous_end, filenames = 0, set()
    for index, (entry, output) in enumerate(zip(entries, outputs), 1):
        if not isinstance(entry, dict) or not isinstance(output, dict):
            raise ValueError("Shared-plan parity requires actual entry and output objects")
        start, end, filename = entry.get("start"), entry.get("end"), entry.get("filename")
        if type(start) is not int or type(end) is not int or not previous_end == start < end <= pages \
                or type(entry.get("sequence")) is not int or entry["sequence"] != index \
                or not isinstance(filename, str) or not filename or filename.casefold() in filenames \
                or output.get("filename") != filename or not isinstance(output.get("range"), list) \
                or len(output["range"]) != 2 or any(type(bound) is not int for bound in output["range"]) \
                or output["range"] != [start, end] \
                or not isinstance(output.get("page_ids"), list) \
                or any(type(page) is not int for page in output["page_ids"]) \
                or output.get("page_ids") != list(range(start + 1, end + 1)):
            raise ValueError("Shared-plan preview order, filenames and physical page identities must match outputs")
        filenames.add(filename.casefold())
        previous_end = end
    if previous_end != pages:
        raise ValueError("Shared-plan writer records omitted final physical pages")
    result = case.get("writer_result")
    if not isinstance(result, dict) or type(result.get("total_pages")) is not int \
            or type(result.get("written_count")) is not int or result["total_pages"] != pages \
            or result["written_count"] != len(entries) or result.get("coverage") != coverage \
            or not isinstance(result.get("coverage"), dict) or result["coverage"].get("complete") is not True \
            or type(result["coverage"].get("covered_pages")) is not int \
            or type(result["coverage"].get("section_count")) is not int \
            or result.get("source_identity") != identity or not isinstance(result.get("source_identity"), dict) \
            or type(result["source_identity"].get("size_bytes")) is not int \
            or not isinstance(result.get("mode"), str) or result["mode"] not in PLAN_MODES \
            or "mode" in case and result["mode"] != case["mode"] \
            or not isinstance(result.get("outputs"), list) \
            or any(not isinstance(output, dict) or any(type(output.get(field)) is not int
                                                      for field in ("sequence", "start", "end", "page_count"))
                   for output in result["outputs"]) \
            or [{key: value for key, value in output.items() if key not in {"sha256", "size_bytes"}}
                for output in result["outputs"]] != [{**entry, "page_count": entry["end"] - entry["start"]} for entry in entries]:
        raise ValueError("Shared-plan execution result must preserve the complete preview and source")
    validate_output_execution(result)


def valid_digest(value: object) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(character in "0123456789abcdef" for character in value)


def validate_output_execution(result: dict) -> None:
    """Each current successful writer receipt names a distinct complete run."""
    run_id, final, manifest = result.get("run_id"), result.get("final_directory"), result.get("manifest")
    if type(result.get("schema_version")) is not int or result["schema_version"] != 1 or result.get("status") != "complete" \
            or not isinstance(run_id, str) or len(run_id) != 32 or any(character not in "0123456789abcdef" for character in run_id) \
            or not isinstance(final, str) or not Path(final).is_absolute() \
            or len(Path(final).name.rsplit("_", 1)[-1]) != 32 \
            or any(character not in "0123456789abcdef" for character in Path(final).name.rsplit("_", 1)[-1]) \
            or result.get("manifest_filename") != "WinBookSplit_Manifest.json" \
            or not isinstance(result.get("outputs"), list) or not result["outputs"] \
            or any(not isinstance(item, dict) or not valid_digest(item.get("sha256"))
                   or type(item.get("size_bytes")) is not int or item["size_bytes"] < 1 for item in result["outputs"]) \
            or not isinstance(manifest, dict) or type(manifest.get("schema_version")) is not int \
            or manifest != {"schema_version": 1, "status": "complete", **{key: result.get(key) for key in
                ("run_id", "final_directory", "mode", "total_pages", "source_identity", "coverage", "written_count", "outputs")}}:
        raise ValueError("Current writer evidence requires a unique complete run and matching manifest/digest/size records")


def stable_writer_result(result: dict) -> dict:
    return {key: value for key, value in result.items() if key not in {"run_id", "final_directory", "manifest_filename", "manifest"}}


def validate_output_report(child: dict) -> None:
    observation = child.get("import_observation")
    if type(child.get("schema_version")) is not int or child.get("task_id") != "M2-T01" \
            or child.get("result") != "OUTPUT_TRANSACTION_REGRESSION_PASSED" or child.get("success") is not True \
            or type(child.get("exit_code")) is not int or child["exit_code"] != 0 \
            or child.get("acceptance_ids") != OUTPUT_ACCEPTANCE_IDS \
            or any(child.get(name) is not True for name in ("source_unchanged", "baseline_guards_preserved",
                                                          "input_and_neighbor_unchanged", "owned_temp_removed")) \
            or child.get("immutable_original_commit") != "0de84f367f9bd5ddfa3f408a9c29505d7a39633f" \
            or not isinstance(observation, dict) or observation.get("import_safe") is not True:
        raise ValueError("Output report requires all acceptance, actual publication, immutable guards and preservation promises")
    repeated = validate_plan_case_ids(child.get("repeat_cases"), OUTPUT_REPEAT_IDS, "repeat/same-basename")
    concurrent = validate_plan_case_ids(child.get("concurrent_cases"), OUTPUT_CONCURRENT_IDS, "concurrent output")
    for case in repeated + concurrent:
        validate_plan_parity(case)
        validate_plan_page_content(case)
        if type(case.get("exit_code")) is not int or case["exit_code"] != 0 \
                or case.get("manifest_validated") is not True or case.get("prior_outputs_unchanged") is not True \
                or not isinstance(case.get("stdout"), str) or not isinstance(case.get("stderr"), str) or case["stderr"] \
                or any(output.get("sha256") != written["sha256"] for output, written in
                       zip(case["outputs"], case["writer_result"]["outputs"])):
            raise ValueError("Output success requires actual native/process and preserved complete-manifest evidence")
        frames = [json.loads(line) for line in case["stdout"].splitlines() if line.startswith("{")]
        if len(frames) != 1 or frames[0].get("protocol") != "winbooksplit.result" \
                or frames[0].get("status") != "success" or frames[0].get("execution") != case["writer_result"] \
                or frames[0].get("written_count") != case["written_count"] or frames[0].get("exit_code") != 0:
            raise ValueError("Output CLI frame must identify its exact successful independent writer receipt")
    for cases in (repeated, concurrent):
        if len({case["writer_result"]["run_id"] for case in cases}) != len(cases) \
                or len({Path(case["writer_result"]["final_directory"]) for case in cases}) != len(cases):
            raise ValueError("Output runs require independent unique identities and final directories")
    if any(type(case.get("pid")) is not int or case["pid"] < 1 or case.get("fixed_timestamp") != "20261009-120000"
           or case.get("barrier_synchronized") is not True or not isinstance(case.get("command"), list)
           or not case["command"] or not isinstance(case.get("started_at"), str) or not case["started_at"]
           or not isinstance(case.get("completed_at"), str) or not case["completed_at"] for case in concurrent) \
            or len({case["pid"] for case in concurrent}) != 4:
        raise ValueError("Concurrency requires four actual barrier-synchronized processes under the identical timestamp")
    failures = validate_plan_case_ids(child.get("failure_cases"), set(OUTPUT_FAILURE_CODES), "output failure")
    for case in failures:
        result, record, owner = case.get("result"), case.get("failure_record"), case.get("failure_owner")
        if not isinstance(result, dict) or type(result.get("exit_code")) is not int or result["exit_code"] != 1 \
                or result.get("code") != OUTPUT_FAILURE_CODES[case["id"]] or result.get("status") not in {"write_error", "error"} \
                or type(result.get("written_count")) is not int or result["written_count"] != 0 or result.get("execution", "missing") is not None \
                or type(case.get("exit_code")) is not int or case["exit_code"] != 1 \
                or type(case.get("successful_final_count")) is not int or case["successful_final_count"] != 0 \
                or any(case.get(name) is not True for name in ("cleanup_complete", "neighbor_unchanged", "source_unchanged")) \
                or not isinstance(record, dict) or not isinstance(owner, dict):
            raise ValueError("Injected output failures must retain their nonzero category without any successful final")
        diagnostic = result.get("diagnostic")
        if not isinstance(diagnostic, dict) or diagnostic.get("cleanup_complete") is not True \
                or diagnostic.get("retained_staging", "missing") is not None or diagnostic.get("cleanup_error", "missing") is not None \
                or not isinstance(diagnostic.get("run_id"), str) or len(diagnostic["run_id"]) != 32 \
                or any(character not in "0123456789abcdef" for character in diagnostic["run_id"]) \
                or not isinstance(diagnostic.get("record_path"), str) or not Path(diagnostic["record_path"]).is_absolute() \
                or Path(diagnostic["record_path"]).name != "failure.json" \
                or Path(diagnostic["record_path"]).parent.name != ".WinBookSplit-failed-" + diagnostic["run_id"] \
                or owner != {"schema_version": 1, "kind": "failed", "run_id": diagnostic["run_id"]} \
                or record.get("schema_version") != 1 or record.get("status") != "failed" \
                or record.get("run_id") != diagnostic["run_id"] or record.get("code") != result["code"] \
                or record.get("cleanup_complete") is not True or record.get("retained_staging", "missing") is not None \
                or not isinstance(record.get("message"), str) or not 0 < len(record["message"]) <= 2048 \
                or case["id"] == "mid-write" and (type(case.get("completed_slices_before_failure")) is not int
                                                   or case["completed_slices_before_failure"] != 1):
            raise ValueError("Output failure evidence requires removed owned staging and a bounded separately marked record")
    cleanup = validate_plan_case_ids(child.get("cleanup_cases"), OUTPUT_CLEANUP_IDS, "owned cleanup")
    for case in cleanup:
        rejected = case["id"] not in {"ordinary-owned", "manifest-path-not-authority"}
        if any(case.get(name) is not True for name in ("cleaned", "owned_stage_removed", "outside_sentinel_unchanged", "actual_windows")) \
                or case.get("rejected") is not rejected \
                or rejected and (not isinstance(case.get("error_code"), str) or not case["error_code"]):
            raise ValueError("Cleanup evidence must reject escapes/tampering and preserve all outside sentinels")
        if case["id"] in {"junction-base", "junction-child"}:
            native = case.get("native_junction")
            if not isinstance(native, dict) or type(native.get("exit_code")) is not int or native["exit_code"] != 0 \
                    or native.get("reparse_verified") is not True or not isinstance(native.get("command"), list) or not native["command"]:
                raise ValueError("Ownership acceptance requires both actual controlled Windows junction observations")


def validate_plan_page_content(case: dict) -> None:
    original_content, actual_content = case.get("original_page_content_sha256"), case.get("page_content_sha256")
    if not isinstance(original_content, list) or len(original_content) != case["total_pages"] \
            or any(not isinstance(digest, str) or len(digest) != 64
                   or any(character not in "0123456789abcdef" for character in digest) for digest in original_content) \
            or not isinstance(actual_content, list) or actual_content != original_content:
        raise ValueError("Shared-plan source outputs must preserve every captured original page's synthetic content")


def path_units(value: str) -> int:
    return len(value.encode("utf-16-le")) // 2


def validate_paths_success(case: dict) -> None:
    validate_plan_parity(case)
    validate_plan_page_content(case)
    result = case["writer_result"]
    observation = case.get("no_replanning_observation")
    if observation not in {"direct-api-planner-and-source-reopen-traps", "supported-by-direct-api-and-shared-plan-controls"} \
            or observation == "direct-api-planner-and-source-reopen-traps" and any(
                type(case.get(field)) is not int or case[field] != 0
                for field in ("planner_calls_during_execute", "source_reader_calls_during_execute")):
        raise ValueError("Path parity must distinguish observed direct-API traps from inherited child-plan controls")
    if case.get("manifest_validated") is not True or case.get("input_unchanged") is not True \
            or case.get("owned_outputs_removed") is not True \
            or not isinstance(case.get("input_resolved"), str) \
            or not Path(case["input_resolved"]).is_absolute() \
            or os.path.normcase(case["input_resolved"]) != os.path.normcase(result["source_identity"]["path"]) \
            or not valid_digest(case.get("input_sha256")) or case["input_sha256"] != result["source_identity"]["sha256"] \
            or not isinstance(case.get("output_base"), str) or not Path(case["output_base"]).is_absolute() \
            or Path(result["final_directory"]).parent != Path(case["output_base"]):
        raise ValueError("Path success requires literal source binding, explicit owned publication and cleanup")
    final = Path(result["final_directory"])
    stage = Path(case["output_base"]) / (".WinBookSplit-stage-" + result["run_id"])
    actual_units = [path_units(str(final / item["filename"])) for item in case["outputs"]]
    staged_units = [path_units(str(stage / item["filename"])) for item in case["outputs"]]
    if case.get("final_path_utf16_units") != actual_units or case.get("stage_path_utf16_units") != staged_units \
            or any(type(value) is not int or value > 259 for value in actual_units + staged_units) \
            or any(path_units(str(directory)) > 247 for directory in (final, stage)):
        raise ValueError("Path success requires independently measured staged and published Windows path budgets")
    for item in case["outputs"]:
        name = item["filename"]
        if re.search(r'[<>:"/\\|?*\x00-\x1f\x7f]', name) or name != name.rstrip(" .") \
                or path_units(name) > 255 or any(0xd800 <= ord(character) <= 0xdfff for character in name):
            raise ValueError("Path outputs must remain safe Unicode basenames")


def validate_paths_report(child: dict, requested_shells: list[str]) -> None:
    if type(child.get("schema_version")) is not int or child.get("schema_version") != 1 \
            or child.get("task_id") != "M2-T02" or child.get("result") != "PATH_REGRESSION_PASSED" \
            or child.get("success") is not True or type(child.get("exit_code")) is not int or child["exit_code"] != 0 \
            or child.get("acceptance_ids") != PATH_ACCEPTANCE_IDS \
            or any(child.get(field) is not True for field in ("source_unchanged", "baseline_guards_preserved",
                                                             "input_and_neighbor_unchanged", "owned_temp_removed",
                                                             "machine_settings_unchanged")) \
            or child.get("immutable_original_commit") != "0de84f367f9bd5ddfa3f408a9c29505d7a39633f" \
            or not isinstance(child.get("import_observation"), dict) or child["import_observation"].get("import_safe") is not True:
        raise ValueError("Path report requires the complete M2-T02 acceptance and preservation promises")
    requested = {os.path.normcase(path) for path in requested_shells}
    hosts = validate_plan_case_ids(child.get("host_cases"), {"PS51", "PS7"}, "actual path hosts")
    if len(requested) != 2 or {os.path.normcase(host.get("shell_executable", "")) for host in hosts} != requested \
            or {host.get("host_major") for host in hosts} != {5, 7}:
        raise ValueError("Path acceptance requires exactly the requested actual PS5.1 and PS7 hosts")
    by_host = {host["id"]: host for host in hosts}
    for host in hosts:
        if type(host.get("host_major")) is not int or not isinstance(host.get("host_version"), str) \
                or host["host_major"] != (5 if host["id"] == "PS51" else 7) \
                or not host["host_version"].startswith("5.1." if host["id"] == "PS51" else "7.") \
                or type(host.get("exit_code")) is not int or host["exit_code"] != 0 \
                or host.get("syntax_checked") != ["WinBookSplit.ps1", "engine/WinBookSplit.Paths.ps1"] \
                or type(host.get("syntax_error_count")) is not int or host["syntax_error_count"] != 0:
            raise ValueError("Path hosts require observed versions and successful production syntax parsing")
        policies = host.get("stored_policies")
        if not isinstance(policies, list) or len(policies) != 4 \
                or any(not isinstance(item, dict) or not isinstance(item.get("policy"), str) or not item["policy"] for item in policies) \
                or {item.get("scope") for item in policies} != {"MachinePolicy", "UserPolicy", "CurrentUser", "LocalMachine"} \
                or host.get("policies_after") != policies:
            raise ValueError("Path hosts require actual unchanged stored policy observations")
    literal = validate_plan_case_ids(child.get("literal_cases"), PATH_LITERAL_IDS, "literal launcher cases")
    for case in literal:
        validate_paths_success(case)
        expected_host = by_host["PS51" if case["id"].startswith(("PS51", "BAT")) else "PS7"]
        argument = case.get("input_argument")
        frame = case.get("engine_record")
        if case.get("actual_process") is not True or type(case.get("exit_code")) is not int or case["exit_code"] != 0 \
                or os.path.normcase(case.get("shell_executable", "")) != os.path.normcase(expected_host["shell_executable"]) \
                or not isinstance(argument, str) or any(token not in argument for token in (" ", "[", "]", "'", "&", "(", ")", "%", "!", "章节")) \
                or not argument.endswith(".PdF") or case.get("wildcard_decoy_unchanged") is not True \
                or case.get("metadata_size_matches") is not True or case.get("literal_expansion_preserved") is not True \
                or case.get("source_read_only_attribute_observed") is not True or case.get("source_attributes_restored") is not True \
                or case.get("no_replanning_observation") != "supported-by-direct-api-and-shared-plan-controls" \
                or not isinstance(case.get("reported_size_display"), str) \
                or case["reported_size_display"].replace(",", ".") != case.get("input_size_display") \
                or not isinstance(frame, dict) or frame.get("protocol") != "winbooksplit.result" \
                or type(frame.get("version")) is not int or frame["version"] != 1 \
                or frame.get("status") != "success" or frame.get("code") != "split_complete" or frame.get("mode") != case["mode"] \
                or type(frame.get("exit_code")) is not int or frame["exit_code"] != 0 \
                or type(frame.get("written_count")) is not int or frame["written_count"] != case["written_count"] \
                or frame.get("execution") != case["writer_result"]:
            raise ValueError("Literal launchers must preserve special characters, exact source metadata and native success")
    rejected = validate_plan_case_ids(child.get("rejection_cases"), PATH_REJECTION_IDS, "literal rejection cases")
    for case in rejected:
        host = by_host[case["id"].split("-", 1)[0]]
        corrupt = case["id"].endswith("corrupt-pdf")
        if case.get("actual_process") is not True or type(case.get("exit_code")) is not int or case["exit_code"] != 1 \
                or os.path.normcase(case.get("shell_executable", "")) != os.path.normcase(host["shell_executable"]) \
                or type(case.get("successful_final_count")) is not int or case["successful_final_count"] != 0 \
                or case.get("outputs") != [] or case.get("input_unchanged") is not True \
                or case.get("neighbor_unchanged") is not True or case.get("owned_outputs_removed") is not True \
                or case.get("engine_called") is not corrupt \
                or corrupt and case.get("error_code") != "unreadable_document" \
                or case["id"].endswith("unreadable-held-file") and case.get("native_lock_verified") is not True:
            raise ValueError("Path rejection requires actual nonzero hosts, parser verification and zero successful output")
    filename = child.get("filename_case")
    if not isinstance(filename, dict) or filename.get("passed") is not True:
        raise ValueError("Path report requires actual safe-title writer evidence")
    validate_paths_success(filename)
    if filename.get("no_replanning_observation") != "direct-api-planner-and-source-reopen-traps":
        raise ValueError("Safe-title writer requires actual direct-API source and planner traps")
    titles = validate_plan_case_ids(filename.get("title_cases"), set(PATH_FILENAME_EXPECTATIONS), "safe title cases")
    if filename["written_count"] != len(PATH_FILENAME_EXPECTATIONS) \
            or [entry["filename"] for entry in filename["outputs"]] != list(PATH_FILENAME_EXPECTATIONS.values()) \
            or any(case.get("filename") != PATH_FILENAME_EXPECTATIONS[case["id"]] for case in titles):
        raise ValueError("Safe title evidence must match independent deterministic fallback and Unicode expectations")
    chapters = validate_plan_case_ids(child.get("chapter_cases"), PLAN_MODES, "120-section modes", key="mode")
    for case in chapters:
        validate_paths_success(case)
        if case.get("no_replanning_observation") != "direct-api-planner-and-source-reopen-traps":
            raise ValueError("120-section writers require actual direct-API source and planner traps")
        names = [item["filename"] for item in case["outputs"]]
        if case.get("section_count") != 120 or type(case.get("section_count")) is not int \
                or case.get("number_width") != 3 or case["written_count"] != 120 or case["total_pages"] != 120 \
                or names != sorted(names) or any(not name.startswith(f"{index:03d} - ") for index, name in enumerate(names, 1)):
            raise ValueError("All modes must write 120 lexically chronological complete sections")
    long_case = child.get("long_destination_case")
    if not isinstance(long_case, dict) or long_case.get("passed") is not True:
        raise ValueError("Path report requires an actual destination-bound long-base writer")
    validate_paths_success(long_case)
    if long_case.get("no_replanning_observation") != "direct-api-planner-and-source-reopen-traps":
        raise ValueError("Long destination writer requires actual direct-API source and planner traps")
    naming = long_case.get("output_naming")
    if not isinstance(naming, dict) or naming.get("resolved_base") != long_case["output_base"] \
            or type(naming.get("filename_budget")) is not int or not 1 <= naming["filename_budget"] < 59 \
            or not isinstance(naming.get("run_stem"), str) or not naming["run_stem"] \
            or long_case.get("preview_unchanged") is not True or long_case.get("shortened") is not True \
            or any(path_units(item["filename"]) > naming["filename_budget"] for item in long_case["outputs"]) \
            or naming["filename_budget"] != min(255, 259 - max(path_units(str(Path(long_case["output_base"]) / (".WinBookSplit-stage-" + "0" * 32))),
                                                               path_units(long_case["writer_result"]["final_directory"])) - 1):
        raise ValueError("Long destination evidence must bind a shortened immutable preview within the measured budget")
    failures = validate_plan_case_ids(child.get("destination_failure_cases"), set(PATH_DESTINATION_FAILURE_CODES), "destination rejection cases")
    for case in failures:
        if case.get("error_code") != PATH_DESTINATION_FAILURE_CODES[case["id"]] \
                or type(case.get("successful_final_count")) is not int or case["successful_final_count"] != 0 \
                or case.get("written_count") != 0 or type(case.get("written_count")) is not int \
                or case.get("outputs") != [] or case.get("cleanup_complete") is not True \
                or case.get("source_unchanged") is not True or case.get("neighbor_unchanged") is not True \
                or not isinstance(case.get("error_message"), str) or not case["error_message"] \
                or case["id"] == "unwritable-held-base" and case.get("native_lock_verified") is not True \
                or case["id"] == "full-target" and (case.get("injected_errno") != 28 or case.get("completed_slices_before_failure") != 1):
            raise ValueError("Destination failure requires an explicit actionable error, zero published files and owned cleanup")


def validate_conversion_success(case: dict, calibre_path: str) -> None:
    result = case.get("writer_result")
    if not isinstance(result, dict) or not isinstance(result.get("manifest"), dict):
        raise ValueError("Conversion requires a complete writer result and manifest")
    ebook_fields = {"original_ebook_identity", "conversion", "retained_intermediate"}
    pdf_result = {key: value for key, value in result.items() if key not in ebook_fields}
    pdf_result["manifest"] = {key: value for key, value in result["manifest"].items() if key not in ebook_fields}
    validate_plan_parity({**case, "writer_result": pdf_result})
    validate_plan_page_content(case)
    original, conversion = case.get("original_ebook_identity"), case.get("conversion")
    fmt = case.get("format")
    if fmt not in {"epub", "azw3"} or case["total_pages"] != (3 if fmt == "epub" else 4) \
            or not isinstance(original, dict) or original.get("binding") != "ebook_snapshot" \
            or not isinstance(original.get("path"), str) or not Path(original["path"]).is_absolute() \
            or Path(original["path"]).suffix.lower() != "." + fmt \
            or not valid_digest(original.get("sha256")) or type(original.get("size_bytes")) is not int or original["size_bytes"] < 1 \
            or case.get("input_sha256") != original["sha256"] or result.get("original_ebook_identity") != original \
            or not isinstance(conversion, dict) or result.get("conversion") != conversion \
            or os.path.normcase(conversion.get("converter_path", "")) != os.path.normcase(calibre_path) \
            or conversion.get("output_profile") != "tablet" or type(conversion.get("exit_code")) is not int or conversion["exit_code"] != 0 \
            or not isinstance(conversion.get("workspace_cleanup"), dict) \
            or conversion["workspace_cleanup"].get("cleanup_complete") is not True \
            or conversion["workspace_cleanup"].get("retained_staging", "missing") is not None:
        raise ValueError("Conversion success requires the actual converter, original ebook identity and cleaned captured PDF")
    generated = conversion.get("generated_pdf_identity")
    source = conversion.get("original_source_identity")
    if not isinstance(generated, dict) or generated.get("binding") != "reader_snapshot" \
            or not isinstance(generated.get("path"), str) or not Path(generated["path"]).is_absolute() \
            or Path(generated["path"]).suffix.lower() != ".pdf" or generated["path"] == original["path"] \
            or generated.get("sha256") != result["source_identity"]["sha256"] \
            or generated.get("size_bytes") != result["source_identity"]["size_bytes"] \
            or generated.get("path") != result["source_identity"].get("path") \
            or Path(generated["path"]).name != "WinBookSplit_Converted.pdf" \
            or Path(generated["path"]).parent.parent != Path(case.get("output_base", "")) \
            or re.fullmatch(r"\.WinBookSplit-stage-[0-9a-f]{32}", Path(generated["path"]).parent.name) is None \
            or type(generated.get("page_count")) is not int or generated["page_count"] != case["total_pages"] \
            or generated.get("page_content_sha256") != case["original_page_content_sha256"] \
            or not isinstance(source, dict) or source.get("binding") != "ebook_snapshot" or source != original:
        raise ValueError("Conversion identities and actual captured PDF page content must remain distinct and agree")
    argv = conversion.get("argv")
    if not isinstance(argv, list) or len(argv) != 5 \
            or os.path.normcase(argv[0]) != os.path.normcase(calibre_path) \
            or os.path.normcase(argv[1]) != os.path.normcase(original["path"]) \
            or os.path.normcase(argv[2]) != os.path.normcase(generated["path"]) \
            or argv[3:] != ["--output-profile", "tablet"]:
        raise ValueError("Conversion report must preserve the exact reviewed executable argument vector")
    for stream in ("stdout", "stderr"):
        tail, count, truncated = conversion.get(stream + "_tail"), conversion.get(stream + "_total_bytes"), conversion.get(stream + "_truncated")
        if not isinstance(tail, str) or type(count) is not int or count < 0 or type(truncated) is not bool:
            raise ValueError("Conversion process evidence requires both bounded stream observations")
    if type(conversion.get("timeout_seconds")) not in {int, float} or conversion["timeout_seconds"] <= 0 \
            or type(conversion.get("elapsed_seconds")) not in {int, float} or conversion["elapsed_seconds"] < 0:
        raise ValueError("Conversion process evidence requires bounded timing")
    keep, retained = case.get("keep_converted_pdf"), result.get("retained_intermediate", "missing")
    if type(keep) is not bool or case.get("retained_intermediate") != retained \
            or result["manifest"].get("original_ebook_identity") != original \
            or result["manifest"].get("conversion") != conversion \
            or result["manifest"].get("retained_intermediate", "missing") != retained:
        raise ValueError("Conversion manifest must preserve original identity and explicit retention policy")
    if keep:
        if retained != {"filename": "WinBookSplit_Converted.pdf", "sha256": generated["sha256"],
                        "size_bytes": generated["size_bytes"], "page_count": case["total_pages"]} \
                or case.get("retained_sha256_verified") is not True or case.get("retained_page_content_verified") is not True:
            raise ValueError("Retained full PDF must independently match the actual captured conversion")
    elif retained is not None or case.get("retained_absent_verified") is not True:
        raise ValueError("Default conversion must leave no retained intermediate")
    if case.get("chapter_markers") != ["WBS-PAGE-001", "WBS-PAGE-002", "WBS-PAGE-003"] \
            or case.get("converted_reference_page_content_sha256") != case["original_page_content_sha256"] \
            or case.get("manifest_validated") is not True or case.get("input_unchanged") is not True \
            or case.get("neighbor_unchanged") is not True or case.get("owned_outputs_removed") is not True \
            or case.get("workspace_absent_verified") is not True:
        raise ValueError("Conversion acceptance requires exact original markers, all physical pages and preserved neighbors")
    expected_ranges = [[0, 1], [1, 2], [2, case["total_pages"]]]
    if [entry.get("range") for entry in case["outputs"]] != expected_ranges \
            or case["written_count"] != 3 or Path(result["final_directory"]).parent != Path(case.get("output_base", "")):
        raise ValueError("Conversion must publish exactly the three requested complete sections in its explicit base")


def validate_conversion_report(child: dict, requested_shells: list[str], calibre_path: str) -> None:
    if type(child.get("schema_version")) is not int or child["schema_version"] != 1 \
            or child.get("task_id") != "M2-T03" or child.get("result") != "CONVERSION_REGRESSION_PASSED" \
            or child.get("success") is not True or type(child.get("exit_code")) is not int or child["exit_code"] != 0 \
            or child.get("acceptance_ids") != CONVERSION_ACCEPTANCE_IDS \
            or any(child.get(field) is not True for field in ("source_unchanged", "baseline_guards_preserved",
                                                             "input_and_neighbor_unchanged", "owned_temp_removed", "machine_settings_unchanged")) \
            or child.get("immutable_original_commit") != "0de84f367f9bd5ddfa3f408a9c29505d7a39633f" \
            or not isinstance(child.get("import_observation"), dict) or child["import_observation"].get("import_safe") is not True:
        raise ValueError("Conversion report requires every M2-T03 acceptance and preservation promise")
    calibre = child.get("calibre")
    if not isinstance(calibre, dict) or os.path.normcase(calibre.get("path", "")) != os.path.normcase(calibre_path) \
            or calibre.get("version") != "9.15.0" or calibre.get("sha256") != CONVERSION_CALIBRE_SHA256 \
            or type(calibre.get("size_bytes")) is not int or calibre["size_bytes"] != 36104 \
            or not isinstance(calibre.get("version_observation"), dict) \
            or calibre["version_observation"].get("exit_code") != 0 \
            or "calibre 9.15.0" not in calibre["version_observation"].get("stdout", ""):
        raise ValueError("Conversion acceptance requires the actual pinned external Calibre executable")
    provenance = child.get("fixture_provenance")
    if not isinstance(provenance, dict) or provenance.get("kind") != "original-offline-ebook-fixtures" \
            or provenance.get("authored_original") is not True or provenance.get("remote_resources") is not False \
            or provenance.get("license") != "MIT" or provenance.get("chapter_markers") != ["WBS-PAGE-001", "WBS-PAGE-002", "WBS-PAGE-003"]:
        raise ValueError("Conversion requires original offline ebook provenance, never private or renamed-format fixtures")
    files = provenance.get("files")
    if not isinstance(files, list) or len(files) != 2 \
            or any(not isinstance(item, dict) or not valid_digest(item.get("sha256"))
                   or type(item.get("size_bytes")) is not int or item["size_bytes"] < 1 for item in files) \
            or {item.get("format") for item in files} != {"epub", "azw3"} \
            or any(item.get("origin") != ("authored-epub" if item["format"] == "epub" else "actual-calibre-conversion") for item in files):
        raise ValueError("Conversion requires both hashed authored EPUB and actual generated AZW3")
    generation = provenance.get("azw3_generation")
    if not isinstance(generation, dict) or type(generation.get("exit_code")) is not int or generation["exit_code"] != 0 \
            or not isinstance(generation.get("argv"), list) or len(generation["argv"]) != 5 \
            or os.path.normcase(generation["argv"][0]) != os.path.normcase(calibre_path) \
            or not generation["argv"][1].endswith(".epub") or not generation["argv"][2].endswith(".azw3") \
            or generation["argv"][3:] != ["--output-profile", "tablet"]:
        raise ValueError("Genuine AZW3 fixture requires the actual pinned converter command and native success")
    references = child.get("reference_conversions")
    if not isinstance(references, dict) or set(references) != {"epub", "azw3"}:
        raise ValueError("Conversion requires independent real PDF reference conversions for both actual formats")
    for fmt, reference in references.items():
        process = reference.get("process") if isinstance(reference, dict) else None
        count = 3 if fmt == "epub" else 4
        if not isinstance(process, dict) or type(process.get("exit_code")) is not int or process["exit_code"] != 0 \
                or not isinstance(process.get("argv"), list) or len(process["argv"]) != 5 \
                or os.path.normcase(process["argv"][0]) != os.path.normcase(calibre_path) \
                or not process["argv"][1].endswith("." + fmt) or not process["argv"][2].endswith(".pdf") \
                or process["argv"][3:] != ["--output-profile", "tablet"] \
                or type(reference.get("page_count")) is not int or reference["page_count"] != count \
                or not valid_digest(reference.get("sha256")) or type(reference.get("size_bytes")) is not int or reference["size_bytes"] < 1 \
                or not isinstance(reference.get("page_content_sha256"), list) or len(reference["page_content_sha256"]) != count \
                or any(not valid_digest(digest) for digest in reference["page_content_sha256"]):
            raise ValueError("Independent real conversion references require exact physical page content and native command evidence")
    hosts = validate_plan_case_ids(child.get("host_cases"), {"PS51", "PS7"}, "conversion actual hosts")
    requested = {os.path.normcase(path) for path in requested_shells}
    if len(requested) != 2 or {os.path.normcase(host.get("shell_executable", "")) for host in hosts} != requested:
        raise ValueError("Conversion requires both requested actual shell hosts")
    by_host = {host["id"]: host for host in hosts}
    for host in hosts:
        if host.get("host_major") != (5 if host["id"] == "PS51" else 7) \
                or not isinstance(host.get("host_version"), str) \
                or not host["host_version"].startswith("5.1." if host["id"] == "PS51" else "7.") \
                or host.get("exit_code") != 0 or host.get("syntax_error_count") != 0 \
                or host.get("syntax_checked") != ["WinBookSplit.ps1", "engine/WinBookSplit.Paths.ps1", "engine/WinBookSplit.Diagnostics.ps1"]:
            raise ValueError("Conversion hosts require actual supported versions and successful syntax parsing")
        policies = host.get("stored_policies")
        if not isinstance(policies, list) or len(policies) != 4 \
                or {item.get("scope") for item in policies if isinstance(item, dict)} != {"MachinePolicy", "UserPolicy", "CurrentUser", "LocalMachine"} \
                or any(not isinstance(item.get("policy"), str) for item in policies) or host.get("policies_after") != policies:
            raise ValueError("Conversion host evidence requires unchanged actual stored policies")
    real = validate_plan_case_ids(child.get("real_cases"), CONVERSION_REAL_IDS, "real conversion/retention")
    for case in real:
        validate_conversion_success(case, calibre_path)
        host_id, fmt, retention = case["id"].split("-")
        frame = case.get("engine_record")
        if case.get("actual_process") is not True or type(case.get("exit_code")) is not int or case["exit_code"] != 0 \
                or os.path.normcase(case.get("shell_executable", "")) != os.path.normcase(by_host[host_id]["shell_executable"]) \
                or case["format"] != fmt or case["keep_converted_pdf"] is not (retention == "keep") \
                or case.get("source_read_only_attribute_observed") is not True or case.get("source_attributes_restored") is not True \
                or case.get("page_content_observation") != "actual-converter-metadata-compared-with-independent-real-conversion-and-slices" \
                or case.get("no_replanning_observation") != "supported-by-direct-api-and-shared-plan-controls" \
                or not isinstance(frame, dict) or frame.get("protocol") != "winbooksplit.result" \
                or type(frame.get("version")) is not int or frame["version"] != 1 \
                or frame.get("status") != "success" or frame.get("code") != "split_complete" \
                or type(frame.get("exit_code")) is not int or frame["exit_code"] != 0 \
                or frame.get("mode") != case["mode"] or type(frame.get("written_count")) is not int \
                or frame["written_count"] != case["written_count"] or frame.get("execution") != case["writer_result"]:
            raise ValueError("Real conversion cases require actual read-only launchers and accurately scoped content evidence")
        if case["original_page_content_sha256"] != references[fmt]["page_content_sha256"]:
            raise ValueError("Actual converted slices must match the independently captured format reference")
    direct = validate_plan_case_ids(child.get("direct_api_cases"), {"epub", "azw3"}, "direct conversion snapshot")
    for case in direct:
        validate_conversion_success(case, calibre_path)
        if case.get("page_content_observation") != "independent-captured-reader-and-real-conversion-and-slices" \
                or case.get("no_replanning_observation") != "direct-api-planner-and-source-reopen-traps" \
                or case.get("captured_pdf_bytes_verified") is not True \
                or any(type(case.get(field)) is not int or case[field] != 0 for field in ("planner_calls_during_execute", "source_reader_calls_during_execute")):
            raise ValueError("Direct ebook execution requires actual captured-reader parity without replanning/reopening")
        if case["original_page_content_sha256"] != references[case["format"]]["page_content_sha256"]:
            raise ValueError("Direct captured reader must match the independent real PDF reference")
    invalid = validate_plan_case_ids(child.get("invalid_converter_cases"), CONVERSION_INVALID_IDS, "zero-exit invalid converters")
    for case in invalid:
        host_id = case["id"].split("-", 1)[0]
        frame = case.get("engine_record")
        converter = case.get("converter_process")
        if case.get("actual_process") is not True or case.get("exit_code") != 1 or case.get("converter_exit_code") != 0 \
                or os.path.normcase(case.get("shell_executable", "")) != os.path.normcase(by_host[host_id]["shell_executable"]) \
                or not isinstance(frame, dict) or frame.get("code") != "conversion_output_invalid" \
                or frame.get("exit_code") != 1 or frame.get("written_count") != 0 or frame.get("execution", "missing") is not None \
                or not isinstance(converter, dict) or type(converter.get("exit_code")) is not int or converter["exit_code"] != 0 \
                or "WBS-FAKE-NATIVE-EXIT-0" not in converter.get("stdout_tail", "") \
                or not isinstance(frame.get("diagnostic"), dict) or frame["diagnostic"].get("conversion") != converter \
                or case.get("outputs") != [] or case.get("successful_final_count") != 0 \
                or case.get("no_success_summary") is not True or case.get("workspace_absent_verified") is not True \
                or any(case.get(field) is not True for field in ("input_unchanged", "neighbor_unchanged", "owned_outputs_removed")):
            raise ValueError("Zero-exit converter controls must reject every invalid PDF without split/success output")
    bat = child.get("bat_missing_converter_case")
    if not isinstance(bat, dict) or bat.get("passed") is not True or bat.get("actual_process") is not True \
            or bat.get("exit_code") != 1 or bat.get("outputs") != [] or bat.get("successful_final_count") != 0 \
            or any(bat.get(field) is not True for field in ("input_unchanged", "neighbor_unchanged", "owned_outputs_removed", "no_success_summary",
                    "preflight_before_console_output", "console_record_absent", "engine_invocation_record_absent",
                    "dependency_success_record_absent", "setup_guidance_verified")) \
            or bat.get("conversion_started") is not False or type(bat.get("written_count")) is not int or bat["written_count"] != 0 \
            or not isinstance(bat.get("dependency_error"), dict) or bat["dependency_error"].get("Code") != "converter_not_found" \
            or not isinstance(bat["dependency_error"].get("Attempts"), list) \
            or any(not isinstance(item, dict) or item.get("Accepted") is not False or item.get("Probe") is not None
                   for item in bat["dependency_error"]["Attempts"]) \
            or bat.get("scope") != "actual-unchanged-BAT-missing-trusted-converter; successful-discovery-separately-covered-by-runtime-acceptance":
        raise ValueError("Conversion evidence must accurately preserve actual BAT missing-converter failure")


def validate_plan_report(child: dict) -> None:
    properties, observation = child.get("seeded_plan_cases"), child.get("import_observation")
    if type(child.get("schema_version")) is not int or child.get("task_id") != "M1-T05" \
            or child.get("result") != "SHARED_PLAN_REGRESSION_PASSED" or child.get("success") is not True \
            or type(child.get("exit_code")) is not int or child["exit_code"] != 0 \
            or child.get("acceptance_ids") != PLAN_ACCEPTANCE_IDS \
            or not isinstance(properties, dict) or type(properties.get("seed")) is not int \
            or properties["seed"] != 20261009 or type(properties.get("count")) is not int \
            or properties["count"] != 300 or properties.get("passed") is not True \
            or not isinstance(properties.get("per_mode"), dict) or set(properties["per_mode"]) != PLAN_MODES \
            or any(type(count) is not int or count != 100 for count in properties["per_mode"].values()) \
            or any(child.get(name) is not True for name in ("source_unchanged", "baseline_guards_preserved",
                                                          "input_and_neighbor_unchanged", "owned_temp_removed")) \
            or child.get("immutable_original_commit") != "0de84f367f9bd5ddfa3f408a9c29505d7a39633f" \
            or not isinstance(observation, dict) or observation.get("import_safe") is not True:
        raise ValueError("Shared-plan report requires all acceptance, seeded mode coverage, immutable guards and preservation promises")
    structural = validate_plan_case_ids(child.get("structural_cases"), PLAN_STRUCTURAL_IDS, "structural")
    for case in structural:
        if case["id"] == "valid-whole-document":
            if case.get("accepted") is not True:
                raise ValueError("Shared-plan report omitted valid whole-document acceptance")
        elif case.get("rejected") is not True or not isinstance(case.get("error_code"), str) or not case["error_code"]:
            raise ValueError("Shared-plan structural rejection requires its actual error code")
    previews = validate_plan_case_ids(child.get("preview_cases"), PLAN_PREVIEW_IDS, "preview")
    for case in previews:
        if type(case.get("chapter_files_written")) is not int or case["chapter_files_written"] != 0 \
                or case["id"] == "read-only-preview" and case.get("mutation_rejected") is not True \
                or case["id"] == "isolated-preview-data" and (case.get("detached_from_input") is not True
                                                              or case.get("unbound_execution_rejected") is not True):
            raise ValueError("Shared-plan preview requires read-only detached metadata and no chapter writes")
    parity = validate_plan_case_ids(child.get("parity_cases"), PLAN_MODES, "mode parity", key="mode")
    for case in parity:
        validate_plan_parity(case)
        validate_plan_page_content(case)
    source_cases = validate_plan_case_ids(child.get("source_cases"), PLAN_SOURCE_IDS, "source binding")
    for case in source_cases:
        identity = case.get("source_identity")
        if case.get("outcome") != "bound_original_preserved" or case.get("replacement_unchanged") is not True \
                or not isinstance(identity, dict) or identity.get("binding") != "reader_snapshot" \
                or not isinstance(identity.get("path"), str) or not identity["path"] \
                or type(identity.get("size_bytes")) is not int or identity["size_bytes"] < 1 \
                or not isinstance(identity.get("sha256"), str) or len(identity["sha256"]) != 64 \
                or any(character not in "0123456789abcdef" for character in identity["sha256"]) \
                or case.get("captured_sha256") != identity["sha256"] \
                or not isinstance(case.get("replacement_sha256"), str) \
                or len(case["replacement_sha256"]) != 64 \
                or any(character not in "0123456789abcdef" for character in case["replacement_sha256"]) \
                or case["replacement_sha256"] == identity["sha256"]:
            raise ValueError("Shared-plan source cases require preserved original snapshot and deliberately different replacement evidence")
        validate_plan_parity(case)
        validate_plan_page_content(case)
        if case["id"] == "repointed-source":
            deleted = case.get("deleted_path_snapshot")
            if not isinstance(deleted, dict):
                raise ValueError("Repointed source evidence requires its deleted-path snapshot execution")
            validate_plan_parity(deleted)
            validate_plan_page_content(deleted)
            if any(deleted[field] != case[field] for field in ("total_pages", "preview_entries", "coverage",
                                                             "written_count", "source_identity",
                                                             "original_page_content_sha256", "page_content_sha256")):
                raise ValueError("Deleted-path execution must retain the same original preview, writer result and source content")
            if stable_writer_result(deleted["writer_result"]) != stable_writer_result(case["writer_result"]) \
                    or deleted["writer_result"]["run_id"] == case["writer_result"]["run_id"] \
                    or deleted["writer_result"]["final_directory"] == case["writer_result"]["final_directory"]:
                raise ValueError("Repeated deleted-path execution requires the same captured plan in a distinct complete run")


def diagnostic_decision(fallbacks: list[str], choice: str) -> tuple[str, str | None]:
    token = choice.strip().upper()
    if not fallbacks:
        return "cancel", None
    if not token:
        return "pending", None
    if token in {"N", "C"}:
        return "cancel", None
    retry = "manual" if token in {"M", "Y"} else "1" if token == "1" else None
    return ("retry", retry) if retry in fallbacks else ("invalid", None)


def validate_diagnostics_report(child: dict, requested_shells: list[str]) -> None:
    observation, cases = child.get("import_observation"), child.get("engine_cases")
    if type(child.get("schema_version")) is not int or child.get("task_id") != "M1-T06" \
            or child.get("result") != "DIAGNOSTIC_REGRESSION_PASSED" or child.get("success") is not True \
            or type(child.get("exit_code")) is not int or child["exit_code"] != 0 \
            or child.get("acceptance_ids") != DIAGNOSTIC_ACCEPTANCE_IDS \
            or type(child.get("engine_case_count")) is not int or child["engine_case_count"] != 23 \
            or any(child.get(name) is not True for name in ("source_unchanged", "baseline_guards_preserved",
                                                          "input_and_neighbor_unchanged", "owned_temp_removed")) \
            or child.get("immutable_original_commit") != "0de84f367f9bd5ddfa3f408a9c29505d7a39633f" \
            or not isinstance(observation, dict) or observation.get("import_safe") is not True:
        raise ValueError("Diagnostics report requires complete acceptance, import and preserved immutable/source/input/cleanup promises")
    validate_plan_case_ids(cases, set(DIAGNOSTIC_EXPECTATIONS), "diagnostic engine")
    by_id = {case["id"]: case for case in cases}
    for case in cases:
        result = case.get("diagnostic")
        mode, status, code, exit_code, fallback = DIAGNOSTIC_EXPECTATIONS[case["id"]]
        if not isinstance(result, dict) or result.get("protocol") != "winbooksplit.result" \
                or type(result.get("version")) is not int or result["version"] != 1 \
                or result.get("mode") != mode or result.get("status") != status or result.get("code") != code \
                or type(case.get("exit_code")) is not int or case["exit_code"] != exit_code \
                or type(result.get("exit_code")) is not int or result["exit_code"] != exit_code \
                or not isinstance(result.get("message"), str) or not result["message"].strip() \
                or not isinstance(result.get("warnings"), list) or result.get("fallback_modes") != fallback \
                or type(result.get("written_count")) is not int or not isinstance(case.get("outputs"), list) \
                or not isinstance(case.get("stdout"), str) or not isinstance(case.get("stderr"), str) \
                or case.get("input_unchanged") is not True or case.get("neighbor_unchanged") is not True \
                or case.get("no_automatic_fallback") is not True:
            raise ValueError("Diagnostic case categories, process exit, choices and preserved input/output evidence must match")
        framed = [json.loads(line) for line in case["stdout"].splitlines() if line.startswith("{")]
        if framed != [result] or "[NO_BOOKMARKS_FOUND]" in case["stdout"]:
            raise ValueError("Each diagnostic process observation requires exactly its one explicit protocol frame")
        if status == "success":
            if not isinstance(result.get("execution"), dict) or result["written_count"] != len(case["outputs"]) or not case["outputs"]:
                raise ValueError("Successful diagnostics require actual nonempty execution outputs")
            execution = result["execution"]
            validate_plan_parity({**case, "source_identity": execution.get("source_identity"),
                                  "writer_result": execution, "coverage": execution.get("coverage"),
                                  "written_count": result["written_count"], "no_replanning_or_reopening": True})
        elif result["written_count"] != 0 or result.get("execution", "missing") is not None or case["outputs"]:
            raise ValueError("Failed diagnostics cannot claim successful or newly created output")
    if by_id["existing-output"].get("existing_outputs_preserved") is not True:
        raise ValueError("Existing output preservation evidence is required")
    if by_id["success-level1"].get("unicode_preserved") is not True \
            or [record.get("filename") for record in by_id["success-level1"]["outputs"]] \
            != ["01 - Front matter.pdf", "02 - 章节 å.pdf", "03 - 次章 é.pdf"]:
        raise ValueError("Actual Unicode bookmark filename/stdout preservation evidence is required")
    shells = child.get("shell_cases")
    requested = {Path(path) for path in requested_shells}
    if len(requested_shells) != 2 or len(requested) != 2 or not isinstance(shells, list) or len(shells) != 2 \
            or any(not isinstance(shell, dict) or not isinstance(shell.get("shell_executable"), str)
                   or type(shell.get("host_major")) is not int
                   or Path(shell["shell_executable"]) not in requested for shell in shells) \
            or {Path(shell["shell_executable"]) for shell in shells} != requested \
            or {shell.get("host_major") for shell in shells} != {5, 7}:
        raise ValueError("Diagnostics require both requested actual PS5.1 and PS7 hosts")
    handler_ids = set(DIAGNOSTIC_EXPECTATIONS) - {"invalid-mode", "missing-arguments"}
    for shell in shells:
        if shell.get("passed") is not True or type(shell.get("exit_code")) is not int or shell["exit_code"] != 0 \
                or type(shell.get("host_major")) is not int or not isinstance(shell.get("host_version"), str) \
                or not shell["host_version"].startswith("5.1." if shell["host_major"] == 5 else "7.") \
                or shell.get("no_automatic_execution") is not True or shell.get("owned_neighbor_unchanged") is not True:
            raise ValueError("Actual diagnostic handlers require successful bounded pure execution and preserved neighbors")
        handlers = validate_plan_case_ids(shell.get("category_cases"), handler_ids, "actual-host category")
        for record in handlers:
            if record.get("parsed_result") != by_id[record["id"]]["diagnostic"]:
                raise ValueError("Actual handler must consume the captured engine result")
            decisions = record.get("decisions")
            if not isinstance(decisions, list) or len(decisions) != len(DIAGNOSTIC_CHOICES):
                raise ValueError("Actual handler requires every explicit fallback decision")
            for choice, decision in zip(DIAGNOSTIC_CHOICES, decisions):
                expected, retry = diagnostic_decision(record["parsed_result"]["fallback_modes"], choice)
                if not isinstance(decision, dict) or decision.get("choice") != choice \
                        or decision.get("decision") != expected or decision.get("retry_mode", "missing") != retry \
                        or decision.get("fallback_modes") != record["parsed_result"]["fallback_modes"] \
                        or not isinstance(decision.get("message"), str) or not decision["message"].strip():
                    raise ValueError("Actual category handler choices/messages must preserve the requested level")
                if record["id"] == "parents-no-level2" and "Level 2" not in decision["message"]:
                    raise ValueError("No-Level2 handler must identify the requested level")
        protocol = validate_plan_case_ids(shell.get("protocol_cases"), DIAGNOSTIC_PROTOCOL_IDS, "protocol rejection")
        if any(record.get("rejected") is not True for record in protocol):
            raise ValueError("Malformed/mismatched protocol receipts must reject")
        arguments, streams = shell.get("native_argument_probe"), shell.get("stream_probe")
        if not isinstance(arguments, dict) or arguments.get("passed") is not True \
                or arguments.get("actual_arguments") != DIAGNOSTIC_NATIVE_ARGUMENTS \
                or type(arguments.get("exit_code")) is not int or arguments["exit_code"] != 0 \
                or not isinstance(streams, dict) or streams.get("passed") is not True \
                or type(streams.get("exit_code")) is not int or streams["exit_code"] != 1 \
                or streams.get("actual_run_function") is not True or streams.get("arguments_preserved") is not True \
                or type(streams.get("stdout_length")) is not int or streams["stdout_length"] != 200000 \
                or type(streams.get("stderr_length")) is not int or streams["stderr_length"] != 200000:
            raise ValueError("Actual-host native argument/dual-stream evidence must be complete: "
                             + repr({"arguments": arguments, "streams": streams}))
        if type(streams.get("actual_stdout_bytes")) is not int or streams["actual_stdout_bytes"] < 200000 \
                or type(streams.get("actual_stderr_bytes")) is not int or streams["actual_stderr_bytes"] != 200000 \
                or streams.get("streams_complete") is not True or streams.get("stdout_truncated") is not True \
                or streams.get("stderr_truncated") is not True \
                or any(type(streams.get(field)) is not int or not 0 < streams[field] <= 65536
                       for field in ("retained_stdout_bytes", "retained_stderr_bytes")) \
                or type(streams.get("log_size_bytes")) is not int or not 0 < streams["log_size_bytes"] < 150000:
            raise ValueError("Actual-host diagnostic flood must record bounded tails, explicit truncation and full byte totals")


def attach_child_report(step: dict, path: Path, kind: str) -> None:
    """A successful process without its promised evidence is a failed step."""
    try:
        data = path.read_bytes()
        child = json.loads(data.decode("utf-8-sig"))
        if not isinstance(child, dict) or child.get("schema_version") != 1:
            raise ValueError("Expected a schema_version=1 JSON object report")
        if kind == "diagnostics":
            validate_diagnostics_report(child, step.get("requested_shell_paths", []))
        elif kind == "paths":
            validate_paths_report(child, step.get("requested_shell_paths", []))
        elif kind == "conversion":
            validate_conversion_report(child, step.get("requested_shell_paths", []), step.get("requested_calibre_path", ""))
        elif kind == "runtime":
            import importlib.util
            spec = importlib.util.spec_from_file_location("wbs_runtime_receipt_validator", ROOT / "tests/runtime/validate_runtime_report.py")
            validator = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(validator)
            validator.validate_runtime_report(child, step.get("requested_shell_paths", []), step.get("requested_calibre_path", ""))
        elif kind == "process":
            import importlib.util
            spec = importlib.util.spec_from_file_location("wbs_process_receipt_validator", ROOT / "tests/process/validate_process_report.py")
            validator = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(validator)
            validator.validate_process_report(child, step.get("requested_shell_paths", []))
        elif kind == "plan":
            validate_plan_report(child)
        elif kind == "shell":
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
        elif kind == "output":
            validate_output_report(child)
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
                    or bookmarks != [] \
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
                                 "seeded coverage checks, twenty-five writer samples, empty current historical "
                                 "comparisons, safe import, preserved guards/source/inputs and owned cleanup")
            validate_level2_launcher_reference(child.get("level2_launcher_reference"))
            if step.get("requested_shell_paths"):
                validate_manual_entrypoints(child, step["requested_shell_paths"])
        elif kind == "bookmarks":
            cases = child.get("engine_cases")
            normalization = child.get("normalization_cases")
            properties = child.get("seeded_level1_cases")
            writer_samples = child.get("sampled_writer_cases")
            import_observation = child.get("import_observation")
            if type(child.get("schema_version")) is not int \
                    or child.get("task_id") != "M1-T03" \
                    or child.get("result") != "LEVEL1_REGRESSION_PASSED" \
                    or child.get("success") is not True \
                    or type(child.get("exit_code")) is not int or child["exit_code"] != 0 \
                    or child.get("acceptance_ids") != LEVEL1_ACCEPTANCE_IDS \
                    or type(child.get("engine_case_count")) is not int or child["engine_case_count"] != 4 \
                    or not isinstance(cases, list) or len(cases) != 4 \
                    or any(not isinstance(case, dict) or not isinstance(case.get("oracle_id"), str)
                           or case.get("passed") is not True for case in cases) \
                    or {case["oracle_id"] for case in cases} != LEVEL1_ORACLE_IDS \
                    or not isinstance(normalization, list) or len(normalization) != 5 \
                    or any(not isinstance(case, dict) or not isinstance(case.get("id"), str)
                           or case.get("passed") is not True for case in normalization) \
                    or {case["id"] for case in normalization} != LEVEL1_NORMALIZATION_IDS \
                    or not isinstance(properties, dict) or type(properties.get("seed")) is not int \
                    or properties["seed"] != 20261009 or type(properties.get("count")) is not int \
                    or properties["count"] != 150 or properties.get("passed") is not True \
                    or not isinstance(writer_samples, list) or len(writer_samples) != 10 \
                    or any(not isinstance(case, dict) or type(case.get("index")) is not int
                           or not isinstance(case.get("outputs"), list) or not case["outputs"]
                           for case in writer_samples) \
                    or {case["index"] for case in writer_samples} != set(range(10)) \
                    or child.get("source_unchanged") is not True \
                    or child.get("baseline_guards_preserved") is not True \
                    or child.get("input_and_neighbor_unchanged") is not True \
                    or child.get("owned_temp_removed") is not True \
                    or child.get("immutable_original_commit") != "0de84f367f9bd5ddfa3f408a9c29505d7a39633f" \
                    or not isinstance(import_observation, dict) or import_observation.get("import_safe") is not True:
                raise ValueError("Level 1 report requires all four targets, five normalization cases, "
                                 "150 seeded plans, ten writer samples, safe import, preserved "
                                 "guards/source/inputs and owned cleanup")
        elif kind == "level2":
            cases = child.get("engine_cases")
            hierarchy = child.get("hierarchy_cases")
            properties = child.get("seeded_level2_cases")
            writer_samples = child.get("sampled_writer_cases")
            import_observation = child.get("import_observation")
            if type(child.get("schema_version")) is not int \
                    or child.get("task_id") != "M1-T04" or child.get("result") != "LEVEL2_REGRESSION_PASSED" \
                    or child.get("success") is not True \
                    or type(child.get("exit_code")) is not int or child["exit_code"] != 0 \
                    or child.get("acceptance_ids") != LEVEL2_ACCEPTANCE_IDS \
                    or type(child.get("engine_case_count")) is not int or child["engine_case_count"] != 4 \
                    or not isinstance(cases, list) or len(cases) != 4 \
                    or any(not isinstance(case, dict) or not isinstance(case.get("oracle_id"), str)
                           or case.get("passed") is not True for case in cases) \
                    or {case["oracle_id"] for case in cases} != LEVEL2_ORACLE_IDS \
                    or not isinstance(hierarchy, list) or len(hierarchy) != 6 \
                    or any(not isinstance(case, dict) or not isinstance(case.get("id"), str)
                           or case.get("passed") is not True for case in hierarchy) \
                    or {case["id"] for case in hierarchy} != LEVEL2_HIERARCHY_IDS \
                    or not isinstance(properties, dict) or type(properties.get("seed")) is not int \
                    or properties["seed"] != 20261009 or type(properties.get("count")) is not int \
                    or properties["count"] != 150 or properties.get("passed") is not True \
                    or not isinstance(writer_samples, list) or len(writer_samples) != 10 \
                    or any(not isinstance(case, dict) or type(case.get("index")) is not int
                           or not isinstance(case.get("outputs"), list) or not case["outputs"]
                           for case in writer_samples) \
                    or {case["index"] for case in writer_samples} != set(range(10)) \
                    or child.get("source_unchanged") is not True or child.get("baseline_guards_preserved") is not True \
                    or child.get("input_and_neighbor_unchanged") is not True or child.get("owned_temp_removed") is not True \
                    or child.get("immutable_original_commit") != "0de84f367f9bd5ddfa3f408a9c29505d7a39633f" \
                    or not isinstance(import_observation, dict) or import_observation.get("import_safe") is not True:
                raise ValueError("Level 2 report requires four targets, six hierarchy cases, 150 seeded plans, "
                                 "ten writer samples, safe import, preserved guards/source/inputs and owned cleanup")
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
                       WBS_TEST_REPORT=str(report), WBS_TEST_WORK=str(work), WBS_TEST_PYTHON=sys.executable,
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
        "conversion_scope": "Actual original offline EPUB/AZW3 conversion and retained/default PDF cases under both requested hosts"
                            if not args.failure_probe and args.layer in {"conversion", "full"}
                            else "NOT_ASSERTED: this layer does not execute the real Calibre conversion acceptance route",
        "not_run": ["interactive preview UI", "in-person console interaction",
                    "human Explorer drag/drop", "unsupported ebooks/remote-resource behavior/DRM",
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
                    "meaning": "Corrected manual targets and corrected six-section Level 2 launcher reference",
                    "requested_shell_paths": [str(shell) for shell in args.shell_path],
                    **run_command(command, work, environment=environment)}
            attach_child_report(step, child_report, "manual")
            steps.append(step)
        if not args.failure_probe and args.layer in {"bookmarks", "full"}:
            child_report = work / "level1-regression.json"
            command = [sys.executable, "-I", "-B",
                       str(ROOT / "tests/bookmarks/characterize_level1.py"),
                       "--report", str(child_report)]
            step = {"name": "level1-regression",
                    "meaning": "Corrected Level 1 target cases and bounded outline normalization",
                    **run_command(command, work, environment=environment)}
            attach_child_report(step, child_report, "bookmarks")
            steps.append(step)
        if not args.failure_probe and args.layer in {"level2", "full"}:
            child_report = work / "level2-regression.json"
            command = [sys.executable, "-I", "-B", str(ROOT / "tests/bookmarks/characterize_level2.py"),
                       "--report", str(child_report)]
            step = {"name": "level2-regression", "meaning": "Corrected parent-aware Level 2 targets and hierarchy checks",
                    **run_command(command, work, environment=environment)}
            attach_child_report(step, child_report, "level2")
            steps.append(step)
        if not args.failure_probe and args.layer in {"plan", "full"}:
            child_report = work / "shared-plan-regression.json"
            command = [sys.executable, "-I", "-B", str(ROOT / "tests/plans/characterize_plan.py"),
                       "--report", str(child_report)]
            step = {"name": "shared-plan-regression",
                    "meaning": "Shared immutable plan validation, preview/writer parity and captured source binding",
                    **run_command(command, work, environment=environment)}
            attach_child_report(step, child_report, "plan")
            steps.append(step)
        if not args.failure_probe and args.layer in {"diagnostics", "full"}:
            child_report = work / "diagnostic-regression.json"
            command = [sys.executable, "-I", "-B", str(ROOT / "tests/diagnostics/characterize_diagnostics.py"),
                       "--report", str(child_report)]
            for shell in args.shell_path:
                command.extend(["--shell-path", str(shell)])
            step = {"name": "diagnostic-regression", "requested_shell_paths": [str(shell) for shell in args.shell_path],
                    "meaning": "Distinct structured diagnostics and actual category/choice handlers under both requested hosts",
                    **run_command(command, work, environment=environment)}
            attach_child_report(step, child_report, "diagnostics")
            steps.append(step)
        if not args.failure_probe and args.layer in {"output", "full"}:
            child_report = work / "output-regression.json"
            command = [sys.executable, "-I", "-B", str(ROOT / "tests/output/characterize_output.py"),
                       "--report", str(child_report)]
            step = {"name": "output-regression", "meaning": "Isolated validated publication, actual fixed-time concurrency and owned Windows cleanup",
                    **run_command(command, work, environment=environment)}
            attach_child_report(step, child_report, "output")
            steps.append(step)
        if not args.failure_probe and args.layer in {"paths", "full"}:
            child_report = work / "path-regression.json"
            command = [sys.executable, "-I", "-B", str(ROOT / "tests/paths/characterize_paths.py"),
                       "--report", str(child_report)]
            for shell in args.shell_path:
                command.extend(["--shell-path", str(shell)])
            step = {"name": "path-regression", "requested_shell_paths": [str(shell) for shell in args.shell_path],
                    "meaning": "Actual literal Windows launchers, safe destination-bound names and conservative path budgets",
                    **run_command(command, work, environment=environment)}
            attach_child_report(step, child_report, "paths")
            steps.append(step)
        if not args.failure_probe and args.layer in {"conversion", "full"}:
            child_report = work / "conversion-regression.json"
            command = [sys.executable, "-I", "-B", str(ROOT / "tests/conversion/characterize_conversion.py"),
                       "--report", str(child_report), "--calibre-path", str(args.calibre_path)]
            for shell in args.shell_path:
                command.extend(["--shell-path", str(shell)])
            step = {"name": "conversion-regression", "requested_shell_paths": [str(shell) for shell in args.shell_path],
                    "requested_calibre_path": str(args.calibre_path),
                    "meaning": "Actual offline EPUB/AZW3 conversion, owned intermediate retention and invalid zero-exit converter rejection",
                    **run_command(command, work, environment=environment)}
            attach_child_report(step, child_report, "conversion")
            steps.append(step)
        if not args.failure_probe and args.layer in {"runtime", "full"}:
            child_report = work / "runtime-regression.json"
            command = [sys.executable, "-I", "-B", str(ROOT / "tests/runtime/characterize_runtime.py"),
                       "--report", str(child_report), "--calibre-path", str(args.calibre_path)]
            for shell in args.shell_path:
                command.extend(["--shell-path", str(shell)])
            step = {"name": "runtime-regression", "requested_shell_paths": [str(shell) for shell in args.shell_path],
                    "requested_calibre_path": str(args.calibre_path),
                    "meaning": "Actual isolated interpreter/import selection, pre-output dependency rejection and trusted converter discovery",
                    **run_command(command, work, environment=environment, timeout=600)}
            attach_child_report(step, child_report, "runtime")
            steps.append(step)
        if not args.failure_probe and args.layer in {"process", "full"}:
            child_report = work / "process-regression.json"
            command = [sys.executable, "-I", "-B", str(ROOT / "tests/process/characterize_process.py"),
                       "--report", str(child_report)]
            for shell in args.shell_path:
                command.extend(["--shell-path", str(shell)])
            step = {"name": "process-regression", "requested_shell_paths": [str(shell) for shell in args.shell_path],
                    "meaning": "Actual owned-tree UTF-8 bounded dual-stream supervision, literal arguments and unchanged BAT status",
                    **run_command(command, work, environment=environment, timeout=600)}
            attach_child_report(step, child_report, "process")
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
    parser.add_argument("--layer", choices=("python", "shell", "baseline", "extraction", "manual", "bookmarks", "level2", "plan", "diagnostics", "output", "paths", "conversion", "runtime", "process", "full"), default="full")
    parser.add_argument("--report", type=Path, help="New absolute JSON file outside checkout")
    parser.add_argument("--tool-root", type=Path, help="Absolute isolated shell module directory")
    parser.add_argument("--calibre-path", type=Path, help="Actual absolute pinned converter for conversion/runtime/full acceptance")
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
        elif args.layer in {"extraction", "manual", "diagnostics", "paths", "conversion", "runtime", "process"} and args.shell_path:
            if any(not shell.is_absolute() or not shell.is_file() for shell in args.shell_path):
                raise ValueError("Every integration shell path must be an existing absolute executable")
            args.shell_path = [shell.resolve() for shell in args.shell_path]
        if args.layer in {"diagnostics", "paths", "conversion", "runtime", "process", "full"} and (len(args.shell_path) != 2 or len(set(args.shell_path)) != 2):
            raise ValueError("Diagnostics, paths, conversion, runtime and process require both explicit distinct supported shell hosts")
        if not args.failure_probe and args.layer in {"conversion", "runtime", "full"}:
            if args.calibre_path is None or not args.calibre_path.is_absolute() or not args.calibre_path.is_file():
                raise ValueError("Conversion/runtime/full requires the existing absolute actual pinned --calibre-path")
            args.calibre_path = args.calibre_path.resolve()
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
