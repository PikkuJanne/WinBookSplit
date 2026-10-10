"""Harness regressions only; these are not corrected-engine acceptance tests."""

import importlib.util
import io
from copy import deepcopy
import json
import os
from pathlib import Path
import sys
import tempfile
import unittest
from types import SimpleNamespace
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[2]
SCRIPT = ROOT / "tests/run_tests.py"
SPEC = importlib.util.spec_from_file_location("wbs_test_runner", SCRIPT)
runner = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(runner)


def corrected_level2_reference():
    ranges = [[0, 2], [2, 3], [3, 6], [6, 8], [8, 10], [10, 12]]
    titles = ["Front matter", "A - Opening pages", "A1", "A2", "B - Opening pages", "B1"]
    return {"oracle_id": "BM-03", "passed": True, "exit_code": 0,
            "expected_ranges": ranges, "titles": titles,
            "outputs": [{"filename": f"{index:02d} - {title}.pdf", "range": pair,
                         "page_ids": list(range(pair[0] + 1, pair[1] + 1))}
                        for index, (title, pair) in enumerate(zip(titles, ranges), 1)]}


def transaction_result(result, number):
    result = deepcopy(result)
    result.update(schema_version=1, status="complete", run_id=f"{number:032x}",
                  final_directory=f"C:/synthetic/output/Book_20261009-120000_{number:032x}",
                  manifest_filename="WinBookSplit_Manifest.json")
    for index, output in enumerate(result["outputs"], 1):
        output.update(sha256=f"{index:064x}", size_bytes=100 + index)
    result["manifest"] = {"schema_version": 1, "status": "complete", **{key: deepcopy(result[key]) for key in
        ("run_id", "final_directory", "mode", "total_pages", "source_identity", "coverage", "written_count", "outputs")}}
    return result


def complete_plan_report():
    entries = [{"sequence": 1, "start": 0, "end": 1, "filename": "01 - Opening.pdf"},
               {"sequence": 2, "start": 1, "end": 3, "filename": "02 - Chapter.pdf"}]
    outputs = [{"filename": "01 - Opening.pdf", "range": [0, 1], "page_ids": [1]},
               {"filename": "02 - Chapter.pdf", "range": [1, 3], "page_ids": [2, 3]}]
    parity = {"passed": True, "total_pages": 3, "preview_entries": entries, "outputs": outputs,
              "coverage": {"complete": True, "covered_pages": 3, "section_count": 2}, "written_count": 2,
              "source_identity": {"path": "C:/synthetic/original.pdf", "sha256": "a" * 64,
                                  "size_bytes": 100, "binding": "reader_snapshot"},
              "neighbor_unchanged": True, "no_replanning_or_reopening": True,
              "original_page_content_sha256": ["c" * 64, "d" * 64, "e" * 64],
              "page_content_sha256": ["c" * 64, "d" * 64, "e" * 64]}
    report = {
        "schema_version": 1, "task_id": "M1-T05", "result": "SHARED_PLAN_REGRESSION_PASSED",
        "success": True, "exit_code": 0, "acceptance_ids": ["AC-027", "AC-028", "AC-029"],
        "structural_cases": [{"id": name, "passed": True, "rejected": True, "error_code": "invalid_plan"}
                             for name in ("gap", "overlap", "empty", "negative", "reversed", "overflow", "zero-pages", "noninteger")]
                            + [{"id": "valid-whole-document", "passed": True, "accepted": True}],
        "preview_cases": [{"id": "read-only-preview", "passed": True, "mutation_rejected": True, "chapter_files_written": 0},
                          {"id": "isolated-preview-data", "passed": True, "detached_from_input": True,
                           "unbound_execution_rejected": True, "chapter_files_written": 0}],
        "parity_cases": [{"mode": mode, **deepcopy(parity)} for mode in ("manual", "1", "2")],
        "source_cases": [{"id": name, **deepcopy(parity), "outcome": "bound_original_preserved",
                          "captured_sha256": "a" * 64, "replacement_sha256": "b" * 64, "replacement_unchanged": True,
                          "original_page_content_sha256": ["c" * 64, "d" * 64, "e" * 64],
                          "page_content_sha256": ["c" * 64, "d" * 64, "e" * 64]}
                         for name in ("changed-same-path", "repointed-source")],
        "seeded_plan_cases": {"seed": 20261009, "count": 300, "per_mode": {"manual": 100, "1": 100, "2": 100}, "passed": True},
        "source_unchanged": True, "baseline_guards_preserved": True, "input_and_neighbor_unchanged": True,
        "owned_temp_removed": True, "immutable_original_commit": "0de84f367f9bd5ddfa3f408a9c29505d7a39633f",
        "import_observation": {"import_safe": True},
    }
    for number, case in enumerate(report["parity_cases"] + report["source_cases"], 1):
        case["writer_result"] = {"mode": case.get("mode", "manual"), "total_pages": 3,
                                 "written_count": 2, "coverage": deepcopy(case["coverage"]),
                                 "source_identity": deepcopy(case["source_identity"]),
                                 "outputs": [{**entry, "page_count": entry["end"] - entry["start"]}
                                             for entry in case["preview_entries"]]}
        case["writer_result"] = transaction_result(case["writer_result"], number)
    repointed = report["source_cases"][1]
    repointed["deleted_path_snapshot"] = {field: deepcopy(repointed[field])
                                         for field in (*parity, "writer_result", "original_page_content_sha256", "page_content_sha256")}
    repointed["deleted_path_snapshot"]["writer_result"] = transaction_result(repointed["writer_result"], 6)
    return report


def complete_diagnostic_report(hosts):
    parities = {case["mode"]: case for case in complete_plan_report()["parity_cases"]}
    unicode_parity = parities["1"]
    unicode_ranges = [(0, 3), (3, 6), (6, 10)]
    names = ["01 - Front matter.pdf", "02 - 章节 å.pdf", "03 - 次章 é.pdf"]
    entries = [{"sequence": index, "start": start, "end": end, "filename": filename}
               for index, ((start, end), filename) in enumerate(zip(unicode_ranges, names), 1)]
    unicode_parity.update(total_pages=10, preview_entries=entries, written_count=3,
                          coverage={"complete": True, "covered_pages": 10, "section_count": 3},
                          outputs=[{"filename": filename, "range": [start, end], "page_ids": list(range(start + 1, end + 1))}
                                   for (start, end), filename in zip(unicode_ranges, names)])
    unicode_parity["writer_result"].update(total_pages=10, written_count=3, coverage=deepcopy(unicode_parity["coverage"]),
        outputs=[{**entry, "page_count": entry["end"] - entry["start"]} for entry in entries])
    unicode_parity["writer_result"] = transaction_result(unicode_parity["writer_result"], 2)
    cases = []
    for identifier, (mode, status, code, exit_code, fallback) in runner.DIAGNOSTIC_EXPECTATIONS.items():
        result = {"protocol": "winbooksplit.result", "version": 1, "mode": mode, "status": status, "code": code,
                  "message": "Synthetic diagnostic message", "exit_code": exit_code, "warnings": [],
                  "fallback_modes": fallback, "written_count": 0, "execution": None}
        case = {"id": identifier, "passed": True, "exit_code": exit_code, "outputs": [], "input_unchanged": True,
                "neighbor_unchanged": True, "no_automatic_fallback": True, "stderr": ""}
        if status == "success":
            parity = deepcopy(parities[mode])
            result.update(written_count=parity["written_count"], execution=parity["writer_result"])
            case.update(total_pages=parity["total_pages"], preview_entries=parity["preview_entries"], outputs=parity["outputs"])
        if identifier == "existing-output":
            case["existing_outputs_preserved"] = True
        if identifier == "success-level1":
            case["unicode_preserved"] = True
        case.update(diagnostic=result, stdout=json.dumps(result) + "\n")
        cases.append(case)
    shells = []
    for index, host in enumerate(hosts):
        categories = []
        for case in cases:
            if case["id"] in {"invalid-mode", "missing-arguments"}:
                continue
            decisions = []
            for choice in runner.DIAGNOSTIC_CHOICES:
                decision, retry = runner.diagnostic_decision(case["diagnostic"]["fallback_modes"], choice)
                decisions.append({"choice": choice, "decision": decision, "retry_mode": retry,
                                  "message": "No usable Level 2 bookmarks" if case["id"] == "parents-no-level2" else "Synthetic handler message",
                                  "fallback_modes": case["diagnostic"]["fallback_modes"]})
            categories.append({"id": case["id"], "passed": True, "parsed_result": deepcopy(case["diagnostic"]), "decisions": decisions})
        shells.append({"shell_executable": host, "passed": True, "exit_code": 0, "host_major": 5 if index == 0 else 7,
                       "host_version": "5.1.1234.1" if index == 0 else "7.6.5", "category_cases": categories,
                       "protocol_cases": [{"id": name, "passed": True, "rejected": True} for name in sorted(runner.DIAGNOSTIC_PROTOCOL_IDS)],
                       "native_argument_probe": {"passed": True, "exit_code": 0, "actual_arguments": runner.DIAGNOSTIC_NATIVE_ARGUMENTS},
                       "stream_probe": {"passed": True, "exit_code": 2, "stdout_length": 200000, "stderr_length": 200000,
                                        "actual_run_function": True, "arguments_preserved": True,
                                        "actual_stdout_bytes": 201000, "actual_stderr_bytes": 200000,
                                        "streams_complete": True, "stdout_truncated": True, "stderr_truncated": True,
                                        "retained_stdout_bytes": 65536, "retained_stderr_bytes": 65536,
                                        "log_size_bytes": 132500},
                       "no_automatic_execution": True, "owned_neighbor_unchanged": True})
    return {"schema_version": 1, "task_id": "M1-T06", "result": "DIAGNOSTIC_REGRESSION_PASSED", "success": True, "exit_code": 0,
            "acceptance_ids": ["AC-030", "AC-031"], "engine_case_count": 23, "engine_cases": cases, "shell_cases": shells,
            "source_unchanged": True, "baseline_guards_preserved": True, "input_and_neighbor_unchanged": True, "owned_temp_removed": True,
            "immutable_original_commit": "0de84f367f9bd5ddfa3f408a9c29505d7a39633f", "import_observation": {"import_safe": True}}


def complete_output_report():
    template = complete_plan_report()["parity_cases"][0]

    def success(identifier, number):
        case = deepcopy(template)
        case.update(id=identifier, manifest_validated=True, prior_outputs_unchanged=True, exit_code=0, stderr="")
        case["writer_result"] = transaction_result(case["writer_result"], number)
        for item, written in zip(case["outputs"], case["writer_result"]["outputs"]):
            item["sha256"] = written["sha256"]
        case["stdout"] = json.dumps({"protocol": "winbooksplit.result", "version": 1, "status": "success",
                                     "exit_code": 0, "written_count": case["written_count"], "execution": case["writer_result"]}) + "\n"
        return case

    repeated = [success(name, 10 + index) for index, name in enumerate(sorted(runner.OUTPUT_REPEAT_IDS))]
    concurrent = [success(name, 20 + index) for index, name in enumerate(sorted(runner.OUTPUT_CONCURRENT_IDS))]
    for index, case in enumerate(concurrent):
        case.update(pid=100 + index, fixed_timestamp="20261009-120000", barrier_synchronized=True,
                    command=["C:/synthetic/python.exe", "-I", "-B", "fixed-timestamp-worker.py"],
                    started_at="2026-10-09T12:00:00+00:00", completed_at="2026-10-09T12:00:01+00:00")
    failures = []
    for index, (identifier, code) in enumerate(sorted(runner.OUTPUT_FAILURE_CODES.items()), 30):
        run_id = f"{index:032x}"
        diagnostic = {"run_id": run_id, "cleanup_complete": True, "retained_staging": None, "cleanup_error": None,
                      "record_path": f"C:/synthetic/output/.WinBookSplit-failed-{run_id}/failure.json"}
        result = {"exit_code": 6, "code": code, "status": "write_error" if identifier == "mid-write" else "error",
                  "written_count": 0, "execution": None, "diagnostic": diagnostic}
        failures.append({"id": identifier, "passed": True, "exit_code": 6, "result": result,
                         "successful_final_count": 0, "cleanup_complete": True, "neighbor_unchanged": True, "source_unchanged": True,
                         "failure_owner": {"schema_version": 1, "kind": "failed", "run_id": run_id},
                         "failure_record": {"schema_version": 1, "status": "failed", "code": code, "run_id": run_id,
                                            "cleanup_complete": True, "retained_staging": None, "message": "Authored failure fixture"},
                         "completed_slices_before_failure": 1 if identifier == "mid-write" else None})
    cleanup = [{"id": identifier, "passed": True, "rejected": identifier not in {"ordinary-owned", "manifest-path-not-authority"},
                "cleaned": True, "owned_stage_removed": True, "outside_sentinel_unchanged": True, "actual_windows": True,
                "error_code": "output_ownership_failed", "native_junction": {"exit_code": 0, "command": ["powershell.exe", "New-Item"],
                                                                                "reparse_verified": True}}
               for identifier in sorted(runner.OUTPUT_CLEANUP_IDS)]
    return {"schema_version": 1, "task_id": "M2-T01", "result": "OUTPUT_TRANSACTION_REGRESSION_PASSED", "success": True, "exit_code": 0,
            "acceptance_ids": runner.OUTPUT_ACCEPTANCE_IDS, "repeat_cases": repeated, "concurrent_cases": concurrent,
            "failure_cases": failures, "cleanup_cases": cleanup, "source_unchanged": True, "baseline_guards_preserved": True,
            "input_and_neighbor_unchanged": True, "owned_temp_removed": True,
            "immutable_original_commit": "0de84f367f9bd5ddfa3f408a9c29505d7a39633f", "import_observation": {"import_safe": True}}


def complete_paths_report(hosts):
    def success(number, names, mode="1"):
        case = deepcopy(complete_plan_report()["parity_cases"][0])
        pages = len(names)
        entries = [{"sequence": i, "start": i - 1, "end": i, "filename": name} for i, name in enumerate(names, 1)]
        source = str(Path("C:/synthetic/book [1] O'Neil & (notes) %WBS_PATHS_EXPAND%!WBS_PATHS_EXPAND! 章节.PdF"))
        case.update(mode=mode, total_pages=pages, preview_entries=entries, written_count=pages,
                    coverage={"complete": True, "covered_pages": pages, "section_count": pages},
                    outputs=[{"filename": name, "range": [i - 1, i], "page_ids": [i], "sha256": f"{i:064x}"}
                             for i, name in enumerate(names, 1)],
                    original_page_content_sha256=[f"{i:064x}" for i in range(1, pages + 1)],
                    page_content_sha256=[f"{i:064x}" for i in range(1, pages + 1)])
        case["source_identity"]["path"] = source
        writer = {"mode": mode, "total_pages": pages, "written_count": pages, "coverage": deepcopy(case["coverage"]),
                  "source_identity": deepcopy(case["source_identity"]), "outputs": [{**entry, "page_count": 1} for entry in entries]}
        case["writer_result"] = transaction_result(writer, number)
        final = Path(case["writer_result"]["final_directory"])
        base = final.parent
        case.update(input_resolved=source, input_sha256=case["source_identity"]["sha256"], input_unchanged=True,
                    output_base=str(base), manifest_validated=True, owned_outputs_removed=True,
                    stage_path_utf16_units=[runner.path_units(str(base / (".WinBookSplit-stage-" + case["writer_result"]["run_id"]) / name)) for name in names],
                    final_path_utf16_units=[runner.path_units(str(final / name)) for name in names])
        case.update(no_replanning_observation="direct-api-planner-and-source-reopen-traps",
                    planner_calls_during_execute=0, source_reader_calls_during_execute=0)
        return case

    host_cases = [{"id": identifier, "passed": True, "shell_executable": host, "host_major": major,
                   "host_version": "5.1.1234.1" if major == 5 else "7.6.5", "exit_code": 0,
                   "syntax_checked": ["WinBookSplit.ps1", "engine/WinBookSplit.Paths.ps1"], "syntax_error_count": 0}
                  for identifier, host, major in zip(("PS51", "PS7"), hosts, (5, 7), strict=True)]
    for host in host_cases:
        host["stored_policies"] = [{"scope": name, "policy": "Undefined"} for name in ("MachinePolicy", "UserPolicy", "CurrentUser", "LocalMachine")]
        host["policies_after"] = deepcopy(host["stored_policies"])
    literal = []
    for i, identifier in enumerate(sorted(runner.PATH_LITERAL_IDS), 1):
        case = success(50 + i, ["01 - Section.pdf"], "manual")
        case.update(id=identifier, actual_process=True, exit_code=0, shell_executable=hosts[1 if identifier.startswith("PS7") else 0],
                    input_argument=case["input_resolved"], wildcard_decoy_unchanged=True, metadata_size_matches=True,
                    literal_expansion_preserved=True, reported_size_display="0.00 MB", input_size_display="0.00 MB")
        case.update(source_read_only_attribute_observed=True, source_attributes_restored=True,
                    no_replanning_observation="supported-by-direct-api-and-shared-plan-controls")
        case["engine_record"] = {"protocol": "winbooksplit.result", "version": 1, "status": "success", "code": "split_complete", "mode": "manual", "exit_code": 0,
                                 "written_count": case["written_count"], "execution": deepcopy(case["writer_result"])}
        literal.append(case)
    rejected = [{"id": name, "passed": True, "actual_process": True, "exit_code": 6 if name.endswith("corrupt-pdf") else 2,
                 "shell_executable": hosts[1 if name.startswith("PS7") else 0], "successful_final_count": 0,
                 "outputs": [], "input_unchanged": True, "neighbor_unchanged": True, "owned_outputs_removed": True,
                 "engine_called": name.endswith("corrupt-pdf"), "error_code": "unreadable_document" if name.endswith("corrupt-pdf") else "literal_preflight_rejected",
                 "native_lock_verified": True}
                for name in sorted(runner.PATH_REJECTION_IDS)]
    filename = success(60, list(runner.PATH_FILENAME_EXPECTATIONS.values()))
    filename["title_cases"] = [{"id": name, "passed": True, "filename": expected} for name, expected in runner.PATH_FILENAME_EXPECTATIONS.items()]
    chapter = [dict(success(70 + i, [f"{n:03d} - Chapter {n}.pdf" for n in range(1, 121)], mode), section_count=120, number_width=3)
               for i, mode in enumerate(("manual", "1", "2"))]
    long_case = success(80, ["01 - Shortened title.pdf"])
    long_base = Path("C:/synthetic") / ("b" * (170 - runner.path_units(str(Path("C:/synthetic"))) - 1))
    final = long_base / ("Book_20261009-120000_" + long_case["writer_result"]["run_id"])
    long_case["writer_result"]["final_directory"] = str(final)
    long_case["writer_result"]["manifest"]["final_directory"] = str(final)
    stage = long_base / (".WinBookSplit-stage-" + long_case["writer_result"]["run_id"])
    long_case.update(output_base=str(long_base), final_path_utf16_units=[runner.path_units(str(final / long_case["outputs"][0]["filename"]))],
                     stage_path_utf16_units=[runner.path_units(str(stage / long_case["outputs"][0]["filename"]))])
    budget = min(255, 259 - max(runner.path_units(str(stage)), runner.path_units(str(final))) - 1)
    long_case.update(output_naming={"resolved_base": long_case["output_base"], "run_stem": "Book", "filename_budget": budget},
                     preview_unchanged=True, shortened=True)
    failures = [{"id": name, "passed": True, "error_code": code, "error_message": "Authored actionable destination failure",
                 "successful_final_count": 0, "written_count": 0, "outputs": [], "cleanup_complete": True,
                 "source_unchanged": True, "neighbor_unchanged": True, "native_lock_verified": True,
                 "injected_errno": 28 if name == "full-target" else None,
                 "completed_slices_before_failure": 1 if name == "full-target" else None}
                for name, code in runner.PATH_DESTINATION_FAILURE_CODES.items()]
    return {"schema_version": 1, "task_id": "M2-T02", "result": "PATH_REGRESSION_PASSED", "success": True, "exit_code": 0,
            "acceptance_ids": runner.PATH_ACCEPTANCE_IDS, "host_cases": host_cases, "literal_cases": literal,
            "rejection_cases": rejected, "filename_case": filename, "chapter_cases": chapter,
            "long_destination_case": long_case, "destination_failure_cases": failures,
            "source_unchanged": True, "baseline_guards_preserved": True, "input_and_neighbor_unchanged": True,
            "owned_temp_removed": True, "machine_settings_unchanged": True,
            "immutable_original_commit": "0de84f367f9bd5ddfa3f408a9c29505d7a39633f", "import_observation": {"import_safe": True}}


def complete_conversion_report(hosts, calibre):
    policies = [{"scope": name, "policy": "Undefined"} for name in ("MachinePolicy", "UserPolicy", "CurrentUser", "LocalMachine")]

    def success(fmt, number, keep):
        pages = 3 if fmt == "epub" else 4
        source = str(Path("C:/synthetic/Book." + fmt))
        generated = str(Path("C:/synthetic/output/.WinBookSplit-stage-" + "f" * 32 + "/WinBookSplit_Converted.pdf"))
        vector = [f"{index:064x}" for index in range(1, pages + 1)]
        entries = [{"sequence": index, "start": start, "end": end, "filename": f"{index:02} - Section {index}.pdf",
                    "title": f"Section {index}", "parent_id": None, "reason": "manual", "warnings": []}
                   for index, (start, end) in enumerate(((0, 1), (1, 2), (2, pages)), 1)]
        identity = {"path": generated, "sha256": "b" * 64, "size_bytes": 2048, "binding": "reader_snapshot"}
        original = {"path": source, "sha256": "a" * 64, "size_bytes": 1024, "binding": "ebook_snapshot"}
        conversion = {"converter_path": calibre, "argv": [calibre, source, generated, "--output-profile", "tablet"],
                      "output_profile": "tablet", "exit_code": 0, "timeout_seconds": 1800, "elapsed_seconds": 1.0,
                      "stdout_tail": "Actual controlled converter", "stderr_tail": "", "stdout_total_bytes": 27,
                      "stderr_total_bytes": 0, "stdout_truncated": False, "stderr_truncated": False,
                      "workspace_cleanup": {"cleanup_complete": True, "retained_staging": None},
                      "original_source_identity": original,
                      "generated_pdf_identity": {**identity, "page_count": pages, "page_content_sha256": vector}}
        retained = {"filename": "WinBookSplit_Converted.pdf", "sha256": "b" * 64, "size_bytes": 2048, "page_count": pages} if keep else None
        coverage = {"complete": True, "covered_pages": pages, "section_count": 3}
        execution = transaction_result({"mode": "manual", "total_pages": pages, "source_identity": identity,
                                       "coverage": coverage, "written_count": 3,
                                       "outputs": [{**entry, "page_count": entry["end"] - entry["start"]} for entry in entries]}, number)
        execution.update(original_ebook_identity=original, conversion=conversion, retained_intermediate=retained)
        execution["manifest"].update(original_ebook_identity=deepcopy(original), conversion=deepcopy(conversion), retained_intermediate=deepcopy(retained))
        return {"passed": True, "format": fmt, "mode": "manual", "total_pages": pages, "preview_entries": entries,
                "outputs": [{"filename": entry["filename"], "range": [entry["start"], entry["end"]],
                             "page_ids": list(range(entry["start"] + 1, entry["end"] + 1))} for entry in entries],
                "coverage": coverage, "source_identity": identity, "original_ebook_identity": original, "conversion": conversion,
                "writer_result": execution, "written_count": 3, "keep_converted_pdf": keep, "retained_intermediate": retained,
                "input_sha256": "a" * 64, "output_base": str(Path(execution["final_directory"]).parent),
                "original_page_content_sha256": vector, "page_content_sha256": vector,
                "converted_reference_page_content_sha256": vector, "chapter_markers": ["WBS-PAGE-001", "WBS-PAGE-002", "WBS-PAGE-003"],
                "manifest_validated": True, "input_unchanged": True, "neighbor_unchanged": True,
                "no_replanning_or_reopening": True, "workspace_absent_verified": True, "owned_outputs_removed": True,
                "retained_sha256_verified": keep, "retained_page_content_verified": keep, "retained_absent_verified": not keep}

    real = []
    for index, identifier in enumerate(sorted(runner.CONVERSION_REAL_IDS), 1):
        host, fmt, retention = identifier.split("-")
        case = success(fmt, index, retention == "keep")
        case.update(id=identifier, actual_process=True, exit_code=0, shell_executable=hosts[0 if host == "PS51" else 1],
                    source_read_only_attribute_observed=True, source_attributes_restored=True,
                    page_content_observation="actual-converter-metadata-compared-with-independent-real-conversion-and-slices",
                    no_replanning_observation="supported-by-direct-api-and-shared-plan-controls")
        case["engine_record"] = {"protocol": "winbooksplit.result", "version": 1, "status": "success", "code": "split_complete",
                                 "exit_code": 0, "mode": "manual", "written_count": 3, "execution": deepcopy(case["writer_result"])}
        real.append(case)
    direct = []
    for number, fmt in enumerate(("epub", "azw3"), 20):
        case = success(fmt, number, fmt == "azw3")
        case.update(id=fmt, page_content_observation="independent-captured-reader-and-real-conversion-and-slices",
                    no_replanning_observation="direct-api-planner-and-source-reopen-traps", planner_calls_during_execute=0, source_reader_calls_during_execute=0,
                    captured_pdf_bytes_verified=True)
        direct.append(case)
    invalid = [{"id": identifier, "passed": True, "actual_process": True, "exit_code": 4, "converter_exit_code": 0,
                "shell_executable": hosts[0 if identifier.startswith("PS51") else 1],
                "engine_record": {"code": "conversion_output_invalid", "exit_code": 4, "written_count": 0, "execution": None},
                "outputs": [], "successful_final_count": 0, "no_success_summary": True, "workspace_absent_verified": True,
                "input_unchanged": True, "neighbor_unchanged": True, "owned_outputs_removed": True}
               for identifier in sorted(runner.CONVERSION_INVALID_IDS)]
    for case in invalid:
        observed = {"exit_code": 0, "stdout_tail": "WBS-FAKE-NATIVE-EXIT-0"}
        case["converter_process"] = observed
        case["engine_record"]["diagnostic"] = {"conversion": deepcopy(observed)}
    provenance = {"kind": "original-offline-ebook-fixtures", "authored_original": True, "remote_resources": False, "license": "MIT",
                  "chapter_markers": ["WBS-PAGE-001", "WBS-PAGE-002", "WBS-PAGE-003"],
                  "files": [{"format": fmt, "sha256": "a" * 64, "size_bytes": 1024,
                             "origin": "authored-epub" if fmt == "epub" else "actual-calibre-conversion"} for fmt in ("epub", "azw3")],
                  "azw3_generation": {"argv": [calibre, "original.epub", "original.azw3", "--output-profile", "tablet"], "exit_code": 0}}
    references = {fmt: {"process": {"argv": [calibre, "original." + fmt, "reference.pdf", "--output-profile", "tablet"], "exit_code": 0},
                        "page_count": 3 if fmt == "epub" else 4, "sha256": "b" * 64, "size_bytes": 2048,
                        "page_content_sha256": [f"{index:064x}" for index in range(1, (3 if fmt == "epub" else 4) + 1)]} for fmt in ("epub", "azw3")}
    return {"schema_version": 1, "task_id": "M2-T03", "result": "CONVERSION_REGRESSION_PASSED", "success": True, "exit_code": 0,
            "acceptance_ids": runner.CONVERSION_ACCEPTANCE_IDS, "calibre": {"path": calibre, "version": "9.15.0",
            "sha256": runner.CONVERSION_CALIBRE_SHA256, "size_bytes": 36104,
            "version_observation": {"exit_code": 0, "stdout": "ebook-convert.exe (calibre 9.15.0)"}},
            "fixture_provenance": provenance, "reference_conversions": references,
            "host_cases": [{"id": host, "passed": True, "shell_executable": hosts[index], "host_major": 5 if host == "PS51" else 7,
                            "host_version": "5.1.26100.9444" if host == "PS51" else "7.6.5", "exit_code": 0, "syntax_error_count": 0,
                            "syntax_checked": ["WinBookSplit.ps1", "engine/WinBookSplit.Paths.ps1", "engine/WinBookSplit.Diagnostics.ps1"],
                            "stored_policies": policies, "policies_after": policies} for index, host in enumerate(("PS51", "PS7"))],
            "real_cases": real, "direct_api_cases": direct, "invalid_converter_cases": invalid,
            "bat_missing_converter_case": {"passed": True, "actual_process": True, "exit_code": 3, "outputs": [], "successful_final_count": 0,
                "input_unchanged": True, "neighbor_unchanged": True, "owned_outputs_removed": True, "no_success_summary": True,
                "preflight_before_console_output": True, "console_record_absent": True, "engine_invocation_record_absent": True,
                "dependency_success_record_absent": True, "conversion_started": False, "written_count": 0, "setup_guidance_verified": True,
                "dependency_error": {"Code": "converter_not_found", "Attempts": []},
                "scope": "actual-unchanged-BAT-missing-trusted-converter; successful-discovery-separately-covered-by-runtime-acceptance"},
            "source_unchanged": True, "baseline_guards_preserved": True, "input_and_neighbor_unchanged": True,
            "owned_temp_removed": True, "machine_settings_unchanged": True,
            "immutable_original_commit": "0de84f367f9bd5ddfa3f408a9c29505d7a39633f", "import_observation": {"import_safe": True}}


class RunnerTests(unittest.TestCase):
    def test_conversion_hashes_nonempty_false_mapping_content_stream(self):
        path = ROOT / "tests/conversion/characterize_conversion.py"
        spec = importlib.util.spec_from_file_location("content_hash_conversion", path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)

        class ContentStream(dict):
            def get_data(self):
                return b"Original nonempty page content"

        stream = ContentStream()
        self.assertFalse(stream)
        reader = SimpleNamespace(pages=[SimpleNamespace(get_contents=lambda: stream)])
        self.assertEqual(module.page_content(reader), [runner.sha256(stream.get_data())])
        self.assertNotEqual(module.page_content(reader), [runner.sha256(b"")])

    def test_conversion_target_forwards_actual_converter_and_both_hosts(self):
        hosts = [Path("C:/trusted/ps51.exe"), Path("C:/trusted/pwsh.exe")]
        calibre = Path("C:/trusted/ebook-convert.exe")
        args = SimpleNamespace(layer="conversion", failure_probe=None, shell_path=hosts, tool_root=None, calibre_path=calibre)
        commands = []

        def command(argv, cwd, **kwargs):
            commands.append(argv)
            return {"exit_code": 0, "stdout": "synthetic-git\n", "stderr": ""}

        with patch.object(runner, "source_manifest", return_value={"synthetic": "hash"}), \
                patch.object(runner, "run_command", side_effect=command), patch.object(runner, "attach_child_report") as attach:
            report = runner.execute(args)
        routed = [argv for argv in commands if str(ROOT / "tests/conversion/characterize_conversion.py") in argv]
        self.assertEqual(len(routed), 1)
        self.assertEqual(routed[0][-6:], ["--calibre-path", str(calibre), "--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
        self.assertEqual([call.args[2] for call in attach.call_args_list], ["conversion"])
        self.assertEqual(report["steps"][0]["requested_calibre_path"], str(calibre))
        self.assertEqual(len(report["steps"]), 1)
        self.assertTrue(report["success"])

    def test_conversion_evidence_requires_real_formats_content_retention_and_invalid_native_success_rejection(self):
        hosts = ["C:/trusted/ps51.exe", "C:/trusted/pwsh.exe"]
        calibre = "C:/trusted/ebook-convert.exe"
        complete = complete_conversion_report(hosts, calibre)
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "conversion.json"
            path.write_text(json.dumps(complete), encoding="utf-8")
            step = {"exit_code": 0, "requested_shell_paths": hosts, "requested_calibre_path": calibre}
            runner.attach_child_report(step, path, "conversion")
            self.assertEqual(step["exit_code"], 0, step.get("evidence_error"))
            defects = [
                ("wrong converter", lambda data: data["calibre"].__setitem__("path", "C:/unexpected/ebook-convert.exe")),
                ("unverified bytes", lambda data: data["calibre"].__setitem__("sha256", "a" * 64)),
                ("renamed azw3", lambda data: data["fixture_provenance"]["files"][1].__setitem__("origin", "renamed-epub")),
                ("missing genuine generator", lambda data: data["fixture_provenance"].pop("azw3_generation")),
                ("private source", lambda data: data["fixture_provenance"].__setitem__("authored_original", False)),
                ("missing reference", lambda data: data["reference_conversions"].pop("azw3")),
                ("unrequested host", lambda data: data["host_cases"][1].__setitem__("shell_executable", "C:/unexpected/pwsh.exe")),
                ("policy changed", lambda data: data["host_cases"][0].__setitem__("policies_after", [])),
                ("missing real case", lambda data: data["real_cases"].pop()),
                ("known PDF confusion", lambda data: data["real_cases"][0]["original_ebook_identity"].__setitem__("binding", "reader_snapshot")),
                ("no cleanup", lambda data: data["real_cases"][0]["conversion"]["workspace_cleanup"].__setitem__("cleanup_complete", False)),
                ("unsafe workspace retention", lambda data: data["real_cases"][0]["conversion"]["workspace_cleanup"].__setitem__("retained_staging", "C:/synthetic/work")),
                ("different source hash", lambda data: data["real_cases"][0]["conversion"]["generated_pdf_identity"].__setitem__("sha256", "c" * 64)),
                ("missing actual source pages", lambda data: data["real_cases"][0]["conversion"]["generated_pdf_identity"].pop("page_content_sha256")),
                ("dropped TOC", lambda data: data["real_cases"][0].__setitem__("page_content_sha256", ["f" * 64])),
                ("native failure marked pass", lambda data: data["real_cases"][0].__setitem__("exit_code", 1)),
                ("missing actual success frame", lambda data: data["real_cases"][0].pop("engine_record")),
                ("read-only unobserved", lambda data: data["real_cases"][0].pop("source_read_only_attribute_observed")),
                ("metadata claimed independent", lambda data: data["real_cases"][0].__setitem__("page_content_observation", "independent-captured-reader-and-real-conversion-and-slices")),
                ("retained bytes unverified", lambda data: data["real_cases"][1].__setitem__("retained_sha256_verified", False)),
                ("missing default absence", lambda data: data["real_cases"][0].pop("retained_absent_verified")),
                ("manifest no conversion", lambda data: data["real_cases"][0]["writer_result"]["manifest"].pop("conversion")),
                ("missing direct snapshot", lambda data: data["direct_api_cases"].pop()),
                ("replan occurred", lambda data: data["direct_api_cases"][0].__setitem__("planner_calls_during_execute", 1)),
                ("captured bytes unverified", lambda data: data["direct_api_cases"][0].pop("captured_pdf_bytes_verified")),
                ("no invalid-output evidence", lambda data: data["invalid_converter_cases"].pop()),
                ("not zero exit", lambda data: data["invalid_converter_cases"][0].__setitem__("converter_exit_code", 1)),
                ("parent did not observe zero", lambda data: data["invalid_converter_cases"][0].pop("converter_process")),
                ("lost native context", lambda data: data["invalid_converter_cases"][0]["engine_record"].pop("diagnostic")),
                ("invalid output accepted", lambda data: data["invalid_converter_cases"][0]["engine_record"].__setitem__("written_count", 1)),
                ("success summary on failure", lambda data: data["invalid_converter_cases"][0].__setitem__("no_success_summary", False)),
                ("BAT zero failure", lambda data: data["bat_missing_converter_case"].__setitem__("exit_code", 0)),
                ("BAT late failure", lambda data: data["bat_missing_converter_case"].__setitem__("preflight_before_console_output", False)),
                ("BAT console written", lambda data: data["bat_missing_converter_case"].__setitem__("console_record_absent", False)),
                ("BAT conversion started", lambda data: data["bat_missing_converter_case"].__setitem__("conversion_started", True)),
                ("BAT wrong dependency", lambda data: data["bat_missing_converter_case"]["dependency_error"].__setitem__("Code", "runtime_invalid")),
                ("BAT unproved probe", lambda data: data["bat_missing_converter_case"]["dependency_error"].__setitem__("Attempts", [{"Accepted": True, "Probe": None}])),
                ("unpreserved baseline", lambda data: data.__setitem__("baseline_guards_preserved", False)),
            ]
            for name, defect in defects:
                with self.subTest(defect=name):
                    child = deepcopy(complete)
                    defect(child)
                    path.write_text(json.dumps(child), encoding="utf-8")
                    step = {"exit_code": 0, "requested_shell_paths": hosts, "requested_calibre_path": calibre}
                    runner.attach_child_report(step, path, "conversion")
                    self.assertEqual(step["exit_code"], 126)
                    self.assertIn("evidence_error", step)

    def test_paths_target_runs_one_stage_and_forwards_both_actual_hosts(self):
        hosts = [Path("C:/trusted/ps51.exe"), Path("C:/trusted/pwsh.exe")]
        args = SimpleNamespace(layer="paths", failure_probe=None, shell_path=hosts, tool_root=None)
        commands = []

        def command(argv, cwd, **kwargs):
            commands.append(argv)
            return {"exit_code": 0, "stdout": "synthetic-git-identity\n", "stderr": ""}

        with patch.object(runner, "source_manifest", return_value={"synthetic-source": "hash"}), \
                patch.object(runner, "run_command", side_effect=command), patch.object(runner, "attach_child_report") as attach:
            report = runner.execute(args)
        self.assertEqual([step["name"] for step in report["steps"]], ["path-regression"])
        self.assertEqual([call.args[2] for call in attach.call_args_list], ["paths"])
        self.assertEqual(report["steps"][0]["requested_shell_paths"], list(map(str, hosts)))
        self.assertEqual(commands[-1][-4:], ["--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
        self.assertTrue(report["success"])

    def test_paths_evidence_requires_actual_hosts_literal_names_full_coverage_and_failure_cleanup(self):
        hosts = ["C:/trusted/ps51.exe", "C:/trusted/pwsh.exe"]
        complete = complete_paths_report(hosts)
        with tempfile.TemporaryDirectory(prefix="wbs-path-evidence-") as directory:
            path = Path(directory) / "paths.json"
            path.write_text(json.dumps(complete), encoding="utf-8")
            step = {"exit_code": 0, "requested_shell_paths": hosts}
            runner.attach_child_report(step, path, "paths")
            self.assertEqual(step["exit_code"], 0)
            defects = []

            def changed(name, mutate):
                child = deepcopy(complete)
                mutate(child)
                defects.append((name, child))

            for field, value in (("schema_version", True), ("task_id", "M2-T01"), ("result", "wrong"), ("exit_code", False),
                                 ("acceptance_ids", ["AC-036"]), ("owned_temp_removed", False), ("source_unchanged", False),
                                 ("input_and_neighbor_unchanged", False), ("machine_settings_unchanged", False), ("import_observation", {})):
                defects.append((field, {**complete, field: value}))
            for group in ("host_cases", "literal_cases", "rejection_cases", "chapter_cases", "destination_failure_cases"):
                defects.append(("missing " + group, {**complete, group: []}))
                changed("duplicate " + group, lambda child, group=group: child[group].__setitem__(-1, deepcopy(child[group][0])))
            changed("unrequested host", lambda child: child["host_cases"][0].update(shell_executable="C:/unrequested/ps51.exe"))
            changed("missing new helper syntax", lambda child: child["host_cases"][0].update(syntax_checked=["WinBookSplit.ps1"]))
            changed("omitted policy observation", lambda child: child["host_cases"][0].pop("policies_after"))
            changed("wildcard input selected", lambda child: child["literal_cases"][0].update(input_resolved="C:/synthetic/decoy.pdf"))
            changed("lost literal percent", lambda child: child["literal_cases"][0].update(input_argument="C:/synthetic/plain.PdF"))
            changed("wrong wildcard metadata", lambda child: child["literal_cases"][0].update(reported_size_display="1.00 MB"))
            changed("unverified literal cleanup", lambda child: child["literal_cases"][0].pop("owned_outputs_removed"))
            changed("false native success", lambda child: child["literal_cases"][0].update(exit_code=1))
            changed("omitted actual native frame", lambda child: child["literal_cases"][0].pop("engine_record"))
            changed("boolean native exit", lambda child: child["literal_cases"][0]["engine_record"].update(exit_code=False))
            changed("unobserved read-only source", lambda child: child["literal_cases"][0].pop("source_read_only_attribute_observed"))
            changed("source attributes left changed", lambda child: child["literal_cases"][0].update(source_attributes_restored=False))
            changed("unobserved child planner traps", lambda child: child["literal_cases"][0].update(no_replanning_observation="direct-api-planner-and-source-reopen-traps"))
            changed("omitted actual parser", lambda child: next(row for row in child["rejection_cases"] if row["id"].endswith("corrupt-pdf")).update(engine_called=False))
            changed("directory accepted", lambda child: child["rejection_cases"][0].update(exit_code=0))
            changed("unverified unreadability", lambda child: next(row for row in child["rejection_cases"] if row["id"].endswith("unreadable-held-file")).update(native_lock_verified=False))
            changed("unsafe or unstated title", lambda child: child["filename_case"]["title_cases"][0].update(filename="../escape.pdf"))
            changed("missing fallback proof", lambda child: child["filename_case"].pop("title_cases"))
            changed("wrong dynamic width", lambda child: child["chapter_cases"][0].update(number_width=2))
            changed("lost 120th page", lambda child: child["chapter_cases"][0]["outputs"][-1].update(page_ids=[]))
            changed("missing ordinary content parity", lambda child: child["chapter_cases"][0].pop("page_content_sha256"))
            changed("omitted direct source traps", lambda child: child["chapter_cases"][0].pop("no_replanning_observation"))
            changed("execution replanned", lambda child: child["chapter_cases"][0].update(planner_calls_during_execute=1))
            changed("unmeasured full path", lambda child: child["long_destination_case"].pop("final_path_utf16_units"))
            changed("changed frozen preview", lambda child: child["long_destination_case"].update(preview_unchanged=False))
            changed("title exceeds frozen budget", lambda child: child["long_destination_case"]["output_naming"].update(filename_budget=1))
            changed("missing held access proof", lambda child: next(row for row in child["destination_failure_cases"] if row["id"] == "unwritable-held-base").update(native_lock_verified=False))
            changed("wrong errno simulation", lambda child: next(row for row in child["destination_failure_cases"] if row["id"] == "full-target").update(injected_errno=13))
            changed("no actual partial slice", lambda child: next(row for row in child["destination_failure_cases"] if row["id"] == "full-target").update(completed_slices_before_failure=0))
            changed("failed stage retained", lambda child: child["destination_failure_cases"][0].update(cleanup_complete=False))
            changed("failure announces files", lambda child: child["destination_failure_cases"][0].update(written_count=1))
            for name, child in defects:
                with self.subTest(defect=name):
                    path.write_text(json.dumps(child), encoding="utf-8")
                    step = {"exit_code": 0, "requested_shell_paths": hosts}
                    runner.attach_child_report(step, path, "paths")
                    self.assertEqual(step["exit_code"], 126)
                    self.assertIn("evidence_error", step)

    def test_output_target_runs_one_stage_without_shell_arguments(self):
        args = SimpleNamespace(layer="output", failure_probe=None, shell_path=[Path("C:/unused/pwsh.exe")], tool_root=None)
        commands = []

        def successful_command(argv, cwd, **kwargs):
            commands.append(argv)
            return {"exit_code": 0, "stdout": "synthetic-git-identity\n", "stderr": ""}

        with patch.object(runner, "source_manifest", return_value={"synthetic-source": "hash"}), \
                patch.object(runner, "run_command", side_effect=successful_command), \
                patch.object(runner, "attach_child_report") as attach:
            report = runner.execute(args)
        self.assertEqual([step["name"] for step in report["steps"]], ["output-regression"])
        self.assertEqual([call.args[2] for call in attach.call_args_list], ["output"])
        self.assertEqual(len(commands), 3)
        self.assertNotIn("--shell-path", commands[-1])
        self.assertTrue(report["success"])

    def test_output_receipt_requires_actual_unique_publication_failures_and_junction_guards(self):
        complete = complete_output_report()
        with tempfile.TemporaryDirectory(prefix="wbs-output-evidence-") as directory:
            path = Path(directory) / "output.json"
            path.write_text(json.dumps(complete), encoding="utf-8")
            valid = {"exit_code": 0}
            runner.attach_child_report(valid, path, "output")
            self.assertEqual(valid["exit_code"], 0)
            self.assertEqual(valid["evidence"], complete)
            defects = []

            def changed(name, mutate):
                report = deepcopy(complete)
                mutate(report)
                defects.append((name, report))

            for field, value in (("schema_version", True), ("task_id", "M1-T06"), ("result", "wrong"), ("exit_code", False),
                                 ("acceptance_ids", ["AC-032"]), ("owned_temp_removed", False), ("source_unchanged", False),
                                 ("baseline_guards_preserved", False), ("input_and_neighbor_unchanged", False), ("import_observation", {})):
                defects.append((field, {**complete, field: value}))
            for group in ("repeat_cases", "concurrent_cases", "failure_cases", "cleanup_cases"):
                defects.append(("missing " + group, {**complete, group: []}))
                changed("duplicate " + group, lambda report, group=group: report[group].__setitem__(-1, deepcopy(report[group][0])))
            for field in ("run_id", "final_directory", "manifest_filename", "manifest"):
                changed("missing publication " + field, lambda report, field=field: report["repeat_cases"][0]["writer_result"].pop(field))
            changed("wrong manifest", lambda report: report["repeat_cases"][0]["writer_result"]["manifest"].update(status="failed"))
            changed("bad digest", lambda report: report["repeat_cases"][0]["writer_result"]["outputs"][0].update(sha256="bad"))
            changed("reopened digest differs", lambda report: report["repeat_cases"][0]["outputs"][0].update(sha256="f" * 64))
            changed("boolean size", lambda report: report["repeat_cases"][0]["writer_result"]["outputs"][0].update(size_bytes=True))
            changed("lost physical page", lambda report: report["repeat_cases"][0]["outputs"][0].update(page_ids=[]))
            changed("changed content", lambda report: report["repeat_cases"][0].update(page_content_sha256=["f" * 64] * 3))
            changed("prior run changed", lambda report: report["repeat_cases"][0].update(prior_outputs_unchanged=False))
            changed("unbound native frame", lambda report: report["repeat_cases"][0].update(stdout="{}\n"))
            changed("reused process", lambda report: report["concurrent_cases"][1].update(pid=report["concurrent_cases"][0]["pid"]))
            changed("timestamp not fixed", lambda report: report["concurrent_cases"][0].update(fixed_timestamp="different"))
            changed("unsynchronized launches", lambda report: report["concurrent_cases"][0].update(barrier_synchronized=False))
            changed("failure succeeded", lambda report: report["failure_cases"][0]["result"].update(exit_code=0))
            changed("published failed final", lambda report: report["failure_cases"][0].update(successful_final_count=1))
            changed("lost failure diagnostic", lambda report: report["failure_cases"][0]["result"].pop("diagnostic"))
            changed("retained failed stage", lambda report: report["failure_cases"][0]["result"]["diagnostic"].update(retained_staging="C:/foreign"))
            changed("wrong failure owner", lambda report: report["failure_cases"][0]["failure_owner"].update(kind="run"))
            changed("midwrite before first slice", lambda report: next(case for case in report["failure_cases"] if case["id"] == "mid-write").update(completed_slices_before_failure=0))
            changed("outside sentinel changed", lambda report: report["cleanup_cases"][0].update(outside_sentinel_unchanged=False))
            changed("junction only mocked", lambda report: next(case for case in report["cleanup_cases"] if case["id"] == "junction-child")["native_junction"].update(reparse_verified=False))
            changed("junction native failed", lambda report: next(case for case in report["cleanup_cases"] if case["id"] == "junction-base")["native_junction"].update(exit_code=1))
            for description, payload in defects:
                with self.subTest(defect=description):
                    path.write_text(json.dumps(payload), encoding="utf-8")
                    step = {"exit_code": 0}
                    runner.attach_child_report(step, path, "output")
                    self.assertEqual(step["exit_code"], 126)
                    self.assertIn("evidence_error", step)
                    self.assertNotIn("evidence", step)
                    failed = {"exit_code": 19}
                    runner.attach_child_report(failed, path, "output")
                    self.assertEqual(failed["exit_code"], 19)
    def test_ac009_python_failure_stays_failed_after_success(self):
        with tempfile.TemporaryDirectory(prefix="wbs-runner-probe-") as directory:
            work = Path(directory)
            neighbor = work / "neighbor.txt"
            neighbor.write_bytes(b"unrelated synthetic neighbor")
            report_path = work / "python.json"
            result = runner.run_command([sys.executable, "-I", "-B", str(SCRIPT),
                                         "--layer", "python", "--failure-probe", "python",
                                         "--report", str(report_path)], work)
            self.assertEqual(result["exit_code"], 1, result)
            report = json.loads(report_path.read_text(encoding="utf-8"))
            self.assertFalse(report["success"])
            self.assertEqual(report["steps"][0]["exit_code"], 1)
            self.assertEqual(report["steps"][-1]["exit_code"], 0)
            self.assertIn("Intentional synthetic Python failure", report["steps"][0]["stderr"])
            self.assertTrue(report["owned_work_directory_removed"])
            self.assertTrue(report["source"]["source_unchanged"])
            self.assertEqual(neighbor.read_bytes(), b"unrelated synthetic neighbor")

    @unittest.skipUnless(sys.platform == "win32", "Native cmd.exe probe requires Windows")
    def test_ac009_native_failure_stays_failed_after_success(self):
        with tempfile.TemporaryDirectory(prefix="wbs-native-probe-") as directory:
            path = Path(directory) / "native.json"
            result = runner.run_command([sys.executable, "-I", "-B", str(SCRIPT),
                                         "--layer", "python", "--failure-probe", "native",
                                         "--report", str(path)], Path(directory))
            self.assertEqual(result["exit_code"], 1, result)
            report = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(report["steps"][0]["exit_code"], 23)
            self.assertEqual(report["steps"][-1]["exit_code"], 0)
            self.assertFalse(report["success"])

    def test_ac010_repeat_probe_from_unrelated_directory(self):
        before = runner.source_manifest()
        with tempfile.TemporaryDirectory(prefix="wbs-repeat-probe-") as directory:
            work = Path(directory)
            # A hostile same-name module cannot be imported by -I or our explicit loads.
            (work / "json.py").write_text("raise RuntimeError('CWD module imported')", encoding="utf-8")
            observations = []
            for index in range(2):
                path = work / f"repeat-{index}.json"
                result = runner.run_command([sys.executable, "-I", "-B", str(SCRIPT),
                                             "--layer", "python", "--failure-probe", "python",
                                             "--report", str(path)], work)
                self.assertEqual(result["exit_code"], 1, result)
                report = json.loads(path.read_text(encoding="utf-8"))
                self.assertTrue(report["owned_work_directory_removed"])
                self.assertFalse(Path(report["owned_work_directory"]).exists())
                self.assertTrue(report["synthetic_input_unchanged"])
                observations.append(([step["exit_code"] for step in report["steps"]],
                                     report["source"]["tested_paths_digest"]))
            self.assertEqual(observations[0], observations[1])
            self.assertTrue((work / "json.py").exists())
        self.assertEqual(before, runner.source_manifest())

    def test_drain_both_large_process_streams(self):
        with tempfile.TemporaryDirectory(prefix="wbs-pipes-") as directory:
            result = runner.run_command([sys.executable, "-I", "-B", "-c",
                                         "import sys; sys.stdout.write('o'*200000); "
                                         "sys.stderr.write('e'*200000); sys.exit(19)"],
                                        Path(directory), timeout=15)
            self.assertEqual(result["exit_code"], 19)
            self.assertEqual(len(result["stdout"]), 200000)
            self.assertEqual(len(result["stderr"]), 200000)
            self.assertFalse(result["timed_out"])

    def test_timeout_is_nonzero_and_keeps_partial_streams(self):
        with tempfile.TemporaryDirectory(prefix="wbs-timeout-") as directory:
            result = runner.run_command([sys.executable, "-I", "-B", "-c",
                                         "import time; print('before timeout',flush=True); time.sleep(3)"],
                                        Path(directory), timeout=0.5)
            self.assertEqual(result["exit_code"], 124)
            self.assertTrue(result["timed_out"])
            self.assertIn("before timeout", result["stdout"])

    def test_missing_command_is_nonzero(self):
        with tempfile.TemporaryDirectory(prefix="wbs-missing-") as directory:
            result = runner.run_command([str(Path(directory) / "missing.exe")], Path(directory))
            self.assertEqual(result["exit_code"], 125)
            self.assertTrue(result["stderr"])

    def test_evidence_missing_or_invalid_cannot_mask_success(self):
        with tempfile.TemporaryDirectory(prefix="wbs-evidence-") as directory:
            path = Path(directory) / "missing.json"
            step = {"exit_code": 0}
            runner.attach_child_report(step, path, "shell")
            self.assertEqual(step["exit_code"], 126)
            path.write_text("[]", encoding="utf-8")
            step = {"exit_code": 0}
            runner.attach_child_report(step, path, "shell")
            self.assertEqual(step["exit_code"], 126)
            step = {"exit_code": 23}
            runner.attach_child_report(step, path, "shell")
            self.assertEqual(step["exit_code"], 23)
            path.write_text('{"schema_version":1,"success":false,"exit_code":1}', encoding="utf-8")
            step = {"exit_code": 0}
            runner.attach_child_report(step, path, "shell")
            self.assertEqual(step["exit_code"], 126)
            for payload in ({}, {"schema_version": 1, "success": True, "exit_code": 0}):
                path.write_text(json.dumps(payload), encoding="utf-8")
                step = {"exit_code": 0}
                runner.attach_child_report(step, path, "shell")
                self.assertEqual(step["exit_code"], 126)
            for kind, payload in (
                ("shell", {"schema_version": 1, "success": True, "exit_code": 0,
                           "pester": {"result": "Passed", "total": 7}}),
                ("baseline", {"schema_version": 1, "result": "ORIGINAL_BEHAVIOR_REPRODUCED",
                              "engine_case_count": 14}),
            ):
                path.write_text(json.dumps(payload), encoding="utf-8")
                step = {"exit_code": 0}
                runner.attach_child_report(step, path, kind)
                self.assertEqual(step["exit_code"], 0)
                self.assertIn("evidence_sha256", step)

    def test_extraction_count_without_actual_cases_cannot_mask_success(self):
        with tempfile.TemporaryDirectory(prefix="wbs-extraction-evidence-") as directory:
            path = Path(directory) / "incomplete.json"
            path.write_text(json.dumps({"schema_version": 1, "success": True, "exit_code": 0,
                                        "result": "EXTRACTION_EQUIVALENCE_REPRODUCED", "engine_case_count": 14,
                                        "source_unchanged": True, "import_observation": {"import_safe": True}}),
                            encoding="utf-8")
            step = {"exit_code": 0}
            runner.attach_child_report(step, path, "extraction")
            self.assertEqual(step["exit_code"], 126)

    def test_manual_evidence_requires_complete_target_and_preservation_results(self):
        complete = {
            "schema_version": 1, "task_id": "M1-T02", "result": "MANUAL_REGRESSION_PASSED",
            "success": True, "exit_code": 0, "acceptance_ids": [f"AC-{number:03}" for number in range(13, 19)],
            "engine_case_count": 22,
            "engine_cases": [{"oracle_id": f"MAN-{number:02}", "passed": True} for number in range(1, 23)],
            "extra_cli_cases": [{"id": name, "passed": True} for name in sorted(runner.MANUAL_EXTRA_IDS)],
            "sampled_writer_cases": [{"index": number, "outputs": [{"filename": "synthetic.pdf"}]}
                                     for number in range(25)],
            "historical_bookmark_cases": [], "level2_launcher_reference": corrected_level2_reference(),
            "source_unchanged": True, "baseline_guards_preserved": True,
            "input_and_neighbor_unchanged": True, "owned_temp_removed": True,
            "immutable_original_commit": "0de84f367f9bd5ddfa3f408a9c29505d7a39633f",
            "import_observation": {"import_safe": True},
            "seeded_property_cases": {"count": 250, "passed": True, "seed": 20261009},
        }
        with tempfile.TemporaryDirectory(prefix="wbs-manual-evidence-") as directory:
            path = Path(directory) / "manual.json"
            path.write_text(json.dumps(complete), encoding="utf-8")
            step = {"exit_code": 0}
            runner.attach_child_report(step, path, "manual")
            self.assertEqual(step["exit_code"], 0)
            self.assertEqual(step["evidence"], complete)
            self.assertIn("evidence_sha256", step)
            step = {"exit_code": 0, "requested_shell_paths": ["C:/trusted/ps51.exe", "C:/trusted/pwsh.exe"]}
            runner.attach_child_report(step, path, "manual")
            self.assertEqual(step["exit_code"], 126, "Core-only report cannot pass requested host probes")
            self.assertNotIn("evidence", step)

            defects = {
                "boolean schema version": {"schema_version": True},
                "wrong task": {"task_id": "M1-T01"},
                "wrong result": {"result": "EXTRACTION_EQUIVALENCE_REPRODUCED"},
                "wrong case count": {"engine_case_count": 21},
                "missing cases despite correct count": {"engine_cases": []},
                "duplicate oracle IDs": {"engine_cases": complete["engine_cases"][:-1]
                                         + [{"oracle_id": "MAN-01", "passed": True}]},
                "unknown oracle IDs": {"engine_cases": complete["engine_cases"][:-1]
                                       + [{"oracle_id": "MAN-99", "passed": True}]},
                "invalid oracle ID type": {"engine_cases": complete["engine_cases"][:-1]
                                           + [{"oracle_id": [], "passed": True}]},
                "failed target case": {"engine_cases": complete["engine_cases"][:-1]
                                      + [{"oracle_id": "MAN-22", "passed": False}]},
                "missing extra CLI cases": {"extra_cli_cases": []},
                "duplicate extra CLI IDs": {"extra_cli_cases": complete["extra_cli_cases"][:-1]
                                             + [complete["extra_cli_cases"][0]]},
                "failed extra CLI case": {"extra_cli_cases": complete["extra_cli_cases"][:-1]
                                           + [{**complete["extra_cli_cases"][-1], "passed": False}]},
                "missing writer samples": {"sampled_writer_cases": []},
                "duplicate writer indexes": {"sampled_writer_cases": complete["sampled_writer_cases"][:-1]
                                               + [complete["sampled_writer_cases"][0]]},
                "boolean writer index": {"sampled_writer_cases": complete["sampled_writer_cases"][:-1]
                                          + [{"index": True, "outputs": [{"filename": "synthetic.pdf"}]}]},
                "zero writer files": {"sampled_writer_cases": complete["sampled_writer_cases"][:-1]
                                       + [{"index": 24, "outputs": []}]},
                "boolean exit code": {"exit_code": False},
                "incomplete acceptance": {"acceptance_ids": ["AC-013"]},
                "missing current historical declaration": {"historical_bookmark_cases": None},
                "obsolete Level 1 known-bad comparisons": {"historical_bookmark_cases": [
                    {"oracle_id": f"BM-{number:02}", "equivalent": True} for number in range(1, 4)]},
                "obsolete Level 2 known-bad comparison": {"historical_bookmark_cases": [{"oracle_id": "BM-03", "equivalent": True}]},
                "changed bookmark behavior": {"historical_bookmark_cases": complete["historical_bookmark_cases"][:-1]
                                             + [{"oracle_id": "BM-03", "equivalent": False}]},
                "changed source": {"source_unchanged": False},
                "changed original guards": {"baseline_guards_preserved": False},
                "changed inputs": {"input_and_neighbor_unchanged": False},
                "missing temp cleanup": {"owned_temp_removed": False},
                "wrong immutable reference": {"immutable_original_commit": "88c2149b3b3034fbd0d7ef23c2382f4b01648e4c"},
                "unsafe import": {"import_observation": {"import_safe": False}},
                "invalid import observation": {"import_observation": []},
                "unrun properties": {"seeded_property_cases": {"count": 0, "passed": True, "seed": 20261009}},
                "failed properties": {"seeded_property_cases": {"count": 250, "passed": False, "seed": 20261009}},
                "missing reproducible seed": {"seeded_property_cases": {"count": 250, "passed": True}},
                "missing corrected launcher reference": {"level2_launcher_reference": None},
                "wrong corrected title": {"level2_launcher_reference": {**corrected_level2_reference(), "titles": ["A1", "A2", "B1"]}},
                "old crossing reference ranges": {"level2_launcher_reference": {**corrected_level2_reference(),
                                                                               "expected_ranges": [[3, 6], [6, 10], [10, 12]]}},
                "incomplete corrected outputs": {"level2_launcher_reference": {**corrected_level2_reference(),
                                                                                "outputs": corrected_level2_reference()["outputs"][:-1]}},
                "wrong corrected page identity": {"level2_launcher_reference": {**corrected_level2_reference(),
                    "outputs": [{**corrected_level2_reference()["outputs"][0], "page_ids": [2, 1]},
                                *corrected_level2_reference()["outputs"][1:]]}},
                "wrong corrected output title": {"level2_launcher_reference": {**corrected_level2_reference(),
                    "outputs": [{**corrected_level2_reference()["outputs"][0], "filename": "01 - A1.pdf"},
                                *corrected_level2_reference()["outputs"][1:]]}},
            }
            for description, changes in defects.items():
                with self.subTest(defect=description):
                    path.write_text(json.dumps({**complete, **changes}), encoding="utf-8")
                    step = {"exit_code": 0}
                    runner.attach_child_report(step, path, "manual")
                    self.assertEqual(step["exit_code"], 126)
                    self.assertIn("evidence_error", step)
                    self.assertNotIn("evidence", step)
                    # Missing evidence must preserve a real nonzero child exit.
                    step = {"exit_code": 19}
                    runner.attach_child_report(step, path, "manual")
                    self.assertEqual(step["exit_code"], 19)

    def test_requested_manual_entrypoints_require_records_preservation_and_requested_hosts(self):
        hosts = ["C:/trusted/ps51.exe", "C:/trusted/pwsh.exe"]
        probes = [{"id": name, "exit_code": 0, "input_unchanged": True,
                   "cwd_engine_untouched": True, "owned_neighbor_unchanged": True,
                   "shell_executable": hosts[1] if "PS7" in name else hosts[0],
                   "outputs": [{"filename": "synthetic.pdf", "page_ids": [1, 2]}]}
                  for name in sorted(runner.MANUAL_ENTRYPOINT_IDS)]
        complete = {"probes": probes, "probe_count": 6, "parallel_launch_count": 3,
                    "shared_temp_engine_sentinel_unchanged": True, "owned_document_outputs_removed": True}
        runner.validate_manual_entrypoints({"entrypoints": complete}, hosts)
        malformed = {
            "count without records": {"probes": []},
            "duplicate IDs": {"probes": probes[:-1] + [probes[0]]},
            "failed child": {"probes": probes[:-1] + [{**probes[-1], "exit_code": 19}]},
            "boolean child exit": {"probes": probes[:-1] + [{**probes[-1], "exit_code": False}]},
            "changed input": {"probes": probes[:-1] + [{**probes[-1], "input_unchanged": False}]},
            "unrequested host": {"probes": probes[:-1] + [{**probes[-1], "shell_executable": "C:/other/pwsh.exe"}]},
            "only one host used": {"probes": [{**probe, "shell_executable": hosts[0]} for probe in probes]},
            "zero files": {"probes": probes[:-1] + [{**probes[-1], "outputs": []}]},
            "missing parallel coverage": {"parallel_launch_count": 0},
            "changed shared sentinel": {"shared_temp_engine_sentinel_unchanged": False},
            "missing cleanup": {"owned_document_outputs_removed": False},
        }
        for description, changes in malformed.items():
            with self.subTest(defect=description):
                with self.assertRaises(ValueError):
                    runner.validate_manual_entrypoints({"entrypoints": {**complete, **changes}}, hosts)
        with self.assertRaises(ValueError):
            runner.validate_manual_entrypoints({}, hosts)

    @unittest.skipUnless(sys.platform == "win32", "Windows host paths use case-insensitive comparison")
    def test_requested_manual_entrypoints_accept_windows_path_case_variants(self):
        hosts = [r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
                 r"C:\Program Files\PowerShell\7\pwsh.exe"]
        receipts = [r"C:\WINDOWS\System32\WindowsPowerShell\v1.0\powershell.exe",
                    "c:/PROGRAM FILES/PowerShell/7/pwsh.exe"]
        probes = [{"id": name, "exit_code": 0, "input_unchanged": True,
                   "cwd_engine_untouched": True, "owned_neighbor_unchanged": True,
                   "shell_executable": receipts[1] if "PS7" in name else receipts[0],
                   "outputs": [{"filename": "synthetic.pdf", "page_ids": [1, 2]}]}
                  for name in sorted(runner.MANUAL_ENTRYPOINT_IDS)]
        complete = {"probes": probes, "probe_count": 6, "parallel_launch_count": 3,
                    "shared_temp_engine_sentinel_unchanged": True, "owned_document_outputs_removed": True}
        runner.validate_manual_entrypoints({"entrypoints": complete}, hosts)
        # Case normalization must not accept a different executable/location.
        unrequested = [{**probe, "shell_executable": r"C:\WINDOWS\System32\other.exe"}
                       if probe["id"] == "PS51-unrelated" else probe for probe in probes]
        with self.assertRaises(ValueError):
            runner.validate_manual_entrypoints({"entrypoints": {**complete, "probes": unrequested}}, hosts)
        # Repeating the same path with different casing is still only one host.
        with self.assertRaises(ValueError):
            runner.validate_manual_entrypoints({"entrypoints": complete}, [hosts[0], receipts[0]])

    def test_full_selects_all_current_regressions_and_forwards_integration_hosts_and_converter(self):
        hosts = [Path("C:/trusted/ps51.exe"), Path("C:/trusted/pwsh.exe")]
        args = SimpleNamespace(layer="full", failure_probe=None, shell_path=hosts,
                               tool_root=Path("C:/trusted/tool-root"), calibre_path=Path("C:/trusted/ebook-convert.exe"),
                               renderer_path=Path("C:/trusted/pdftoppm.exe"),
                               secondary_python_path=Path("C:/trusted/renderer-python.exe"),
                               report=Path("C:/trusted/full.json"))
        commands = []

        def successful_command(argv, cwd, **kwargs):
            commands.append(argv)
            return {"exit_code": 0, "stdout": "synthetic-git-identity\n", "stderr": ""}

        with patch.object(runner, "source_manifest", return_value={"synthetic-test-source": "hash"}), \
                patch.object(runner, "run_command", side_effect=successful_command), \
                patch.object(runner, "attach_child_report") as attach:
            report = runner.execute(args)
        manual = [command for command in commands
                  if str(ROOT / "tests/manual/characterize_manual.py") in command]
        self.assertEqual(len(manual), 1)
        self.assertEqual(manual[0][-4:], ["--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
        self.assertFalse(any(str(ROOT / "tests/extraction/characterize_extraction.py") in command
                             for command in commands))
        bookmarks = [command for command in commands
                     if str(ROOT / "tests/bookmarks/characterize_level1.py") in command]
        self.assertEqual(len(bookmarks), 1)
        self.assertNotIn("--shell-path", bookmarks[0])
        level2 = [command for command in commands if str(ROOT / "tests/bookmarks/characterize_level2.py") in command]
        self.assertEqual(len(level2), 1)
        self.assertNotIn("--shell-path", level2[0])
        plans = [command for command in commands if str(ROOT / "tests/plans/characterize_plan.py") in command]
        self.assertEqual(len(plans), 1)
        self.assertNotIn("--shell-path", plans[0])
        diagnostics = [command for command in commands if str(ROOT / "tests/diagnostics/characterize_diagnostics.py") in command]
        self.assertEqual(len(diagnostics), 1)
        self.assertEqual(diagnostics[0][-4:], ["--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
        output = [command for command in commands if str(ROOT / "tests/output/characterize_output.py") in command]
        self.assertEqual(len(output), 1)
        self.assertNotIn("--shell-path", output[0])
        paths = [command for command in commands if str(ROOT / "tests/paths/characterize_paths.py") in command]
        self.assertEqual(len(paths), 1)
        self.assertEqual(paths[0][-4:], ["--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
        conversion = [command for command in commands if str(ROOT / "tests/conversion/characterize_conversion.py") in command]
        self.assertEqual(len(conversion), 1)
        self.assertEqual(conversion[0][-6:], ["--calibre-path", str(args.calibre_path), "--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
        runtime = [command for command in commands if str(ROOT / "tests/runtime/characterize_runtime.py") in command]
        self.assertEqual(len(runtime), 1)
        self.assertEqual(runtime[0][-6:], ["--calibre-path", str(args.calibre_path), "--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
        process = [command for command in commands if str(ROOT / "tests/process/characterize_process.py") in command]
        self.assertEqual(len(process), 1)
        self.assertEqual(process[0][-4:], ["--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
        outcomes = [command for command in commands if str(ROOT / "tests/outcomes/characterize_outcomes.py") in command]
        self.assertEqual(len(outcomes), 1)
        self.assertEqual(outcomes[0][-4:], ["--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
        cli = [command for command in commands if str(ROOT / "tests/cli/characterize_cli.py") in command]
        self.assertEqual(len(cli), 1)
        self.assertEqual(cli[0][-6:], ["--calibre-path", str(args.calibre_path), "--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
        launcher = [command for command in commands if str(ROOT / "tests/launcher/characterize_launcher.py") in command]
        self.assertEqual(len(launcher), 1)
        self.assertEqual(launcher[0][-6:], ["--calibre-path", str(args.calibre_path), "--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
        ux = [command for command in commands if str(ROOT / "tests/ux/characterize_ux.py") in command]
        self.assertEqual(len(ux), 1)
        self.assertEqual(ux[0][-6:], ["--calibre-path", str(args.calibre_path), "--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
        support = [command for command in commands if str(ROOT / "tests/support/characterize_support.py") in command]
        self.assertEqual(len(support), 1)
        self.assertEqual(support[0][-6:], ["--calibre-path", str(args.calibre_path), "--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
        fidelity = [command for command in commands if str(ROOT / "tests/fidelity/characterize_fidelity.py") in command]
        self.assertEqual(len(fidelity), 1)
        for flag, value in (("--renderer-path", args.renderer_path), ("--secondary-python-path", args.secondary_python_path),
                            ("--render-directory", args.report.with_name("full-renders"))):
            self.assertEqual(fidelity[0][fidelity[0].index(flag) + 1], str(value))
        self.assertEqual(fidelity[0][-4:], ["--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
        document_policy = [command for command in commands if str(ROOT / "tests/document_policy/characterize_document_policy.py") in command]
        self.assertEqual(len(document_policy), 1)
        self.assertEqual(document_policy[0][-4:], ["--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
        for flag in ("--calibre-path", "--renderer-path", "--secondary-python-path"):
            self.assertNotIn(flag, document_policy[0])
        faults = [command for command in commands if str(ROOT / "tests/faults/characterize_faults.py") in command]
        self.assertEqual(len(faults), 1)
        self.assertEqual(faults[0][-4:], ["--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
        ebooks = [command for command in commands if str(ROOT / "tests/ebooks/characterize_ebooks.py") in command]
        self.assertEqual(len(ebooks), 1)
        for flag, value in (("--renderer-path", args.renderer_path), ("--calibre-path", args.calibre_path),
                            ("--render-directory", args.report.with_name("full-ebook-renders"))):
            self.assertEqual(ebooks[0][ebooks[0].index(flag) + 1], str(value))
        self.assertEqual(ebooks[0][-4:], ["--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
        self.assertEqual([call.args[2] for call in attach.call_args_list], ["shell", "shell", "manual", "bookmarks", "level2", "plan", "diagnostics", "output", "paths", "conversion", "runtime", "process", "outcomes", "cli", "launcher", "ux", "support", "fidelity", "document-policy", "faults", "ebooks"])
        self.assertEqual([step["name"] for step in report["steps"][-19:]],
                         ["manual-regression", "level1-regression", "level2-regression", "shared-plan-regression", "diagnostic-regression", "output-regression", "path-regression", "conversion-regression", "runtime-regression", "process-regression", "outcome-regression", "cli-regression", "launcher-regression", "ux-regression", "support-regression", "fidelity-regression", "document-policy-regression", "fault-regression", "ebook-regression"])
        self.assertEqual(len(report["steps"]), 22)
        self.assertTrue(report["success"])

    def test_ebooks_target_binds_both_routes_and_preserves_either_native_failure(self):
        hosts = [Path("C:/trusted/ps51.exe"), Path("C:/trusted/pwsh.exe")]
        args = SimpleNamespace(layer="ebooks", failure_probe=None, shell_path=hosts,
                               calibre_path=Path("C:/trusted/ebook-convert.exe"),
                               renderer_path=Path("C:/trusted/pdftoppm.exe"),
                               report=Path("C:/external/ebooks.json"))
        for failure in (None, "conversion", "ebooks"):
            with self.subTest(failure=failure):
                commands = []
                def command(argv, cwd, **kwargs):
                    commands.append(argv)
                    failed = failure is not None and str(ROOT / ("tests/conversion/characterize_conversion.py"
                                     if failure == "conversion" else "tests/ebooks/characterize_ebooks.py")) in argv
                    return {"exit_code": 19 if failed else 0, "stdout": "synthetic identity\n", "stderr": ""}
                with patch.object(runner, "source_manifest", return_value={"synthetic": "hash"}), \
                        patch.object(runner, "run_command", side_effect=command), patch.object(runner, "attach_child_report") as attach:
                    report = runner.execute(args)
                self.assertEqual([step["name"] for step in report["steps"]], ["conversion-regression", "ebook-regression"])
                self.assertEqual([call.args[2] for call in attach.call_args_list], ["conversion", "ebooks"])
                self.assertEqual(report["success"], failure is None)
                self.assertEqual(sum(step["exit_code"] != 0 for step in report["steps"]), int(failure is not None))
                self.assertEqual(report["steps"][-1]["requested_source_sha256"], {"synthetic": "hash"})
                self.assertEqual(report["steps"][-1]["requested_shell_paths"], list(map(str, hosts)))
                ebook_command = [argv for argv in commands if str(ROOT / "tests/ebooks/characterize_ebooks.py") in argv][0]
                self.assertEqual(ebook_command[ebook_command.index("--render-directory") + 1],
                                 str(args.report.with_name("ebooks-ebook-renders")))

    def test_faults_target_runs_each_required_cross_host_route_once_and_keeps_native_failure(self):
        hosts = [Path("C:/trusted/ps51.exe"), Path("C:/trusted/pwsh.exe")]
        args = SimpleNamespace(layer="faults", failure_probe=None, shell_path=hosts)
        kinds = ("paths", "process", "outcomes", "faults")
        expected_names = ["path-regression", "process-regression", "outcome-regression", "fault-regression"]
        for failure in (None, *expected_names):
            with self.subTest(failure=failure):
                commands = []

                def command(argv, cwd, **kwargs):
                    commands.append(argv)
                    failing = failure is not None and any(failure.replace("-regression", "") in item for item in argv)
                    # Paths and outcomes use plural directory/script names.
                    if failure == "path-regression":
                        failing = str(ROOT / "tests/paths/characterize_paths.py") in argv
                    elif failure == "outcome-regression":
                        failing = str(ROOT / "tests/outcomes/characterize_outcomes.py") in argv
                    elif failure == "fault-regression":
                        failing = str(ROOT / "tests/faults/characterize_faults.py") in argv
                    return {"exit_code": 19 if failing else 0, "stdout": "synthetic identity\n", "stderr": ""}

                with patch.object(runner, "source_manifest", return_value={"synthetic": "hash"}), \
                        patch.object(runner, "run_command", side_effect=command), patch.object(runner, "attach_child_report") as attach:
                    report = runner.execute(args)
                self.assertEqual([step["name"] for step in report["steps"]], expected_names)
                self.assertEqual([call.args[2] for call in attach.call_args_list], list(kinds))
                for step in report["steps"]:
                    self.assertEqual(step["requested_shell_paths"], list(map(str, hosts)))
                self.assertEqual(report["steps"][-1]["requested_source_sha256"], {"synthetic": "hash"})
                self.assertEqual(report["success"], failure is None)
                self.assertEqual(sum(step["exit_code"] != 0 for step in report["steps"]), int(failure is not None))
                self.assertEqual(len([argv for argv in commands if "--report" in argv]), 4)

    def test_document_policy_target_forwards_only_both_hosts_and_propagates_native_failure(self):
        hosts = [Path("C:/trusted/ps51.exe"), Path("C:/trusted/pwsh.exe")]
        args = SimpleNamespace(layer="document-policy", failure_probe=None, shell_path=hosts)
        for native_exit in (0, 17):
            with self.subTest(native_exit=native_exit):
                commands = []

                def command(argv, cwd, **kwargs):
                    commands.append(argv)
                    return {"exit_code": native_exit if str(ROOT / "tests/document_policy/characterize_document_policy.py") in argv else 0,
                            "stdout": "synthetic identity\n", "stderr": ""}

                with patch.object(runner, "source_manifest", return_value={"synthetic": "hash"}), \
                        patch.object(runner, "run_command", side_effect=command), patch.object(runner, "attach_child_report") as attach:
                    report = runner.execute(args)
                native = [argv for argv in commands if str(ROOT / "tests/document_policy/characterize_document_policy.py") in argv]
                self.assertEqual(len(native), 1)
                self.assertEqual(native[0][-4:], ["--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
                for flag in ("--calibre-path", "--renderer-path", "--secondary-python-path"):
                    self.assertNotIn(flag, native[0])
                self.assertEqual([step["name"] for step in report["steps"]], ["document-policy-regression"])
                self.assertEqual(report["steps"][0]["requested_shell_paths"], list(map(str, hosts)))
                self.assertEqual([call.args[2] for call in attach.call_args_list], ["document-policy"])
                self.assertEqual(report["success"], native_exit == 0)

    def test_document_policy_readme_and_policy_are_bound_in_source_manifest(self):
        with tempfile.TemporaryDirectory(prefix="wbs-policy-runner-source-") as directory:
            root = Path(directory)
            content = {"tests/document_policy/README.md": b"Authored route\n", "docs/PDF_POLICY.md": b"Declared policy\n"}
            for name, data in content.items():
                target = root / name
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(data)
            with patch.object(runner, "ROOT", root):
                before = runner.source_manifest()
                self.assertEqual(before, {name: runner.sha256(data) for name, data in content.items()})
                (root / "docs/PDF_POLICY.md").write_bytes(b"Changed declared policy\n")
                self.assertNotEqual(runner.manifest_digest(before), runner.manifest_digest(runner.source_manifest()))

    def test_ci_workflow_helpers_and_readme_are_bound_in_source_manifest(self):
        with tempfile.TemporaryDirectory(prefix="wbs-ci-runner-source-") as directory:
            root = Path(directory)
            content = {
                ".github/workflows/windows-ci.yml": b"Authored workflow\n",
                "tools/ci/README.md": b"Declared CI scope\n",
                "tools/ci/bootstrap.py": b"Pinned isolated bootstrap\n",
                "tools/ci/check_package_inputs.py": b"Immutable input checks\n",
            }
            for name, data in content.items():
                target = root / name
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(data)
            with patch.object(runner, "ROOT", root):
                before = runner.source_manifest()
                self.assertEqual(before, {name: runner.sha256(data) for name, data in content.items()})
                for name, data in content.items():
                    with self.subTest(name=name):
                        target = root / name
                        target.write_bytes(data + b"Changed\n")
                        self.assertNotEqual(runner.manifest_digest(before), runner.manifest_digest(runner.source_manifest()))
                        target.write_bytes(data)

    def test_document_policy_main_requires_distinct_existing_hosts_without_converter_or_renderer(self):
        # Only argument validation is exercised; these empty authored files are
        # never launched and do not stand in for actual supported host evidence.
        with tempfile.TemporaryDirectory(prefix="wbs-policy-runner-args-") as directory:
            work = Path(directory).resolve()
            hosts = [work / "ps51.exe", work / "pwsh.exe"]
            for host in hosts:
                host.write_bytes(b"")
            for supplied in ([], [hosts[0]], [hosts[0], hosts[0]], [work / "missing.exe", hosts[1]], hosts):
                with self.subTest(hosts=supplied):
                    destination = work / "report.json"
                    argv = [str(SCRIPT), "--layer", "document-policy", "--report", str(destination)]
                    for host in supplied:
                        argv.extend(["--shell-path", str(host)])
                    synthetic = {"success": True, "steps": [{"exit_code": 0}]}
                    with patch.object(sys, "argv", argv), patch.object(runner, "execute", return_value=synthetic) as execute, \
                            patch.object(sys, "stdout", io.StringIO()), patch.object(sys, "stderr", io.StringIO()):
                        code = runner.main()
                    if supplied == hosts:
                        self.assertEqual(code, 0)
                        self.assertEqual(execute.call_args.args[0].shell_path, hosts)
                        self.assertIsNone(execute.call_args.args[0].calibre_path)
                        self.assertIsNone(execute.call_args.args[0].renderer_path)
                        self.assertIsNone(execute.call_args.args[0].secondary_python_path)
                        self.assertTrue(destination.is_file())
                    else:
                        self.assertEqual(code, 1)
                        execute.assert_not_called()
                        self.assertFalse(destination.exists())

    def test_document_policy_attachment_calls_strict_validator_with_cleanup_required(self):
        # This proves runner delegation/error propagation only. Native receipt
        # content/contradiction tests live with the independent policy validator.
        hosts = ["C:/trusted/ps51.exe", "C:/trusted/pwsh.exe"]
        child = {"schema_version": 1, "success": True, "exit_code": 0}
        calls = []

        def validate(data, requested, *, cleanup_complete):
            calls.append((data, requested, cleanup_complete))
            if data.get("reject"):
                raise ValueError("Synthetic strict rejection")

        validator = SimpleNamespace(validate_document_policy_report=validate)
        spec = SimpleNamespace(loader=SimpleNamespace(exec_module=lambda module: None))
        with tempfile.TemporaryDirectory(prefix="wbs-policy-runner-validator-") as directory:
            path = Path(directory) / "report.json"
            with patch.object(importlib.util, "spec_from_file_location", return_value=spec) as load, \
                    patch.object(importlib.util, "module_from_spec", return_value=validator):
                path.write_text(json.dumps(child), encoding="utf-8")
                step = {"exit_code": 0, "requested_shell_paths": hosts}
                runner.attach_child_report(step, path, "document-policy")
                self.assertEqual(load.call_args.args[1], ROOT / "tests/document_policy/validate_document_policy_report.py")
                self.assertEqual(calls, [(child, hosts, True)])
                self.assertEqual(step["evidence"], child)
                self.assertNotIn("failed_evidence", step)
                for native in (0, 7):
                    rejected = {**child, "reject": True, "cleanup_safe": False}
                    path.write_text(json.dumps(rejected), encoding="utf-8")
                    step = {"exit_code": native, "requested_shell_paths": hosts}
                    runner.attach_child_report(step, path, "document-policy")
                    self.assertEqual(step["exit_code"], 126 if native == 0 else native)
                    self.assertNotIn("evidence", step)
                    self.assertEqual(step["failed_evidence"], rejected)
                    self.assertFalse(step["cleanup_safe"])
                    self.assertIn("evidence_error", step)

    def test_document_policy_absent_malformed_and_partial_receipts_fail_closed(self):
        with tempfile.TemporaryDirectory(prefix="wbs-policy-runner-receipt-") as directory:
            path = Path(directory) / "report.json"
            partial = {"schema_version": 1, "task_id": "M4-T02", "success": False,
                       "exit_code": 1, "cleanup_safe": False, "cases": []}
            for raw in (None, "not JSON", json.dumps(partial)):
                if raw is not None:
                    path.write_text(raw, encoding="utf-8")
                step = {"exit_code": 0, "requested_shell_paths": ["C:/trusted/ps51.exe", "C:/trusted/pwsh.exe"]}
                runner.attach_child_report(step, path, "document-policy")
                self.assertEqual(step["exit_code"], 126)
                self.assertIn("evidence_error", step)
                self.assertNotIn("evidence", step)
                if raw is not None and raw.startswith("{"):
                    self.assertEqual(step["failed_evidence"], partial)
                    self.assertEqual(step["failed_evidence_sha256"], runner.sha256(path.read_bytes()))
                    self.assertFalse(step["cleanup_safe"])

    def test_support_target_forwards_both_hosts_and_real_converter(self):
        hosts = [Path("C:/trusted/ps51.exe"), Path("C:/trusted/pwsh.exe")]
        args = SimpleNamespace(layer="support", failure_probe=None, shell_path=hosts,
                               tool_root=None, calibre_path=Path("C:/trusted/ebook-convert.exe"))
        commands = []
        def command(argv, cwd, **kwargs):
            commands.append(argv)
            return {"exit_code": 0, "stdout": "synthetic identity\n", "stderr": ""}
        with patch.object(runner, "source_manifest", return_value={"synthetic": "hash"}), \
                patch.object(runner, "run_command", side_effect=command), patch.object(runner, "attach_child_report") as attach:
            report = runner.execute(args)
        native = [argv for argv in commands if str(ROOT / "tests/support/characterize_support.py") in argv]
        self.assertEqual(len(native), 1)
        self.assertEqual(native[0][-6:], ["--calibre-path", str(args.calibre_path), "--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
        self.assertEqual([step["name"] for step in report["steps"]], ["support-regression"])
        self.assertEqual([call.args[2] for call in attach.call_args_list], ["support"])
        self.assertTrue(report["success"])

    def test_support_absent_malformed_and_retained_failure_receipts_fail_closed(self):
        with tempfile.TemporaryDirectory(prefix="wbs-support-receipt-") as directory:
            path = Path(directory) / "support.json"
            for raw in (None, "not JSON", json.dumps({"schema_version": 1, "task_id": "M3-T05", "success": False,
                            "exit_code": 1, "cleanup_safe": False, "workspace_retained": "C:/synthetic/retained"})):
                if raw is not None:
                    path.write_text(raw, encoding="utf-8")
                step = {"exit_code": 0, "requested_shell_paths": ["C:/trusted/ps51.exe", "C:/trusted/pwsh.exe"],
                        "requested_calibre_path": "C:/trusted/ebook-convert.exe"}
                runner.attach_child_report(step, path, "support")
                self.assertEqual(step["exit_code"], 126)
                self.assertIn("evidence_error", step)
                if raw is not None and raw.startswith("{"):
                    self.assertFalse(step["cleanup_safe"])
                    self.assertIn("failed_evidence", step)
                    native_failure = {**step, "exit_code": 7}
                    runner.attach_child_report(native_failure, path, "support")
                    self.assertEqual(native_failure["exit_code"], 7)

    def test_diagnostic_evidence_requires_categories_zero_outputs_both_hosts_and_all_decisions(self):
        hosts = ["C:/trusted/ps51.exe", "C:/trusted/pwsh.exe"]
        complete = complete_diagnostic_report(hosts)
        with tempfile.TemporaryDirectory(prefix="wbs-diagnostic-evidence-") as directory:
            path = Path(directory) / "diagnostics.json"
            path.write_text(json.dumps(complete), encoding="utf-8")
            valid = {"exit_code": 0, "requested_shell_paths": hosts}
            runner.attach_child_report(valid, path, "diagnostics")
            self.assertEqual(valid["exit_code"], 0)
            self.assertEqual(valid["evidence"], complete)
            self.assertIn("evidence_sha256", valid)
            defects = []

            def changed(description, mutate):
                report = deepcopy(complete)
                mutate(report)
                defects.append((description, report))

            for field, value in (("schema_version", True), ("task_id", "M1-T05"), ("result", "SHARED_PLAN_REGRESSION_PASSED"),
                                 ("success", False), ("exit_code", False), ("acceptance_ids", ["AC-030"]),
                                 ("engine_case_count", 0), ("source_unchanged", False), ("baseline_guards_preserved", False),
                                 ("input_and_neighbor_unchanged", False), ("owned_temp_removed", False),
                                 ("immutable_original_commit", "unknown"), ("import_observation", []), ("engine_cases", []), ("shell_cases", [])):
                defects.append((field, {**complete, field: value}))
            for field, value in (("id", []), ("id", "unknown"), ("passed", False), ("exit_code", 0),
                                 ("outputs", [{"filename": "unwanted.pdf"}]), ("stdout", "no result"), ("input_unchanged", False),
                                 ("neighbor_unchanged", False), ("no_automatic_fallback", False)):
                changed("engine " + field, lambda report, field=field, value=value: report["engine_cases"][0].update({field: value}))
            for field, value in (("protocol", "wrong"), ("version", True), ("status", "success"), ("code", "no_bookmarks_at_level"),
                                 ("written_count", 1), ("execution", {}), ("fallback_modes", ["1", "manual"]), ("warnings", {})):
                changed("category " + field, lambda report, field=field, value=value: report["engine_cases"][0]["diagnostic"].update({field: value}))
            changed("duplicate engine", lambda report: report["engine_cases"].__setitem__(-1, deepcopy(report["engine_cases"][0])))
            changed("existing output changed", lambda report: next(case for case in report["engine_cases"] if case["id"] == "existing-output").update(existing_outputs_preserved=False))
            changed("unicode not verified", lambda report: next(case for case in report["engine_cases"] if case["id"] == "success-level1").update(unicode_preserved=False))
            for field, value in (("shell_executable", "C:/other/pwsh.exe"), ("host_major", []), ("host_version", "6.0"),
                                 ("passed", False), ("exit_code", 19), ("category_cases", []), ("protocol_cases", []),
                                 ("no_automatic_execution", False), ("owned_neighbor_unchanged", False),
                                 ("native_argument_probe", {}), ("stream_probe", {})):
                changed("host " + field, lambda report, field=field, value=value: report["shell_cases"][0].update({field: value}))
            changed("duplicate host", lambda report: report["shell_cases"].__setitem__(1, deepcopy(report["shell_cases"][0])))
            changed("duplicate category", lambda report: report["shell_cases"][0]["category_cases"].__setitem__(-1, deepcopy(report["shell_cases"][0]["category_cases"][0])))
            changed("unparsed category", lambda report: report["shell_cases"][0]["category_cases"][0].update(parsed_result={}))
            changed("missing decisions", lambda report: report["shell_cases"][0]["category_cases"][0].update(decisions=[]))
            changed("automatic retry", lambda report: report["shell_cases"][0]["category_cases"][0]["decisions"][0].update(decision="retry", retry_mode="manual"))
            changed("wrong level choice", lambda report: report["shell_cases"][0]["category_cases"][0]["decisions"][3].update(decision="retry", retry_mode="1"))
            changed("substring choice accepted", lambda report: report["shell_cases"][0]["category_cases"][0]["decisions"][6].update(decision="retry", retry_mode="manual"))
            changed("misleading Level2 message", lambda report: next(case for case in report["shell_cases"][0]["category_cases"] if case["id"] == "parents-no-level2")["decisions"][0].update(message="No outline bookmarks in PDF"))
            changed("accepted invalid protocol", lambda report: report["shell_cases"][0]["protocol_cases"][0].update(rejected=False))
            changed("changed literal argv", lambda report: report["shell_cases"][0]["native_argument_probe"].update(actual_arguments=["wrong"]))
            changed("lost stderr", lambda report: report["shell_cases"][0]["stream_probe"].update(stderr_length=0))
            changed("unbounded stdout tail", lambda report: report["shell_cases"][0]["stream_probe"].update(retained_stdout_bytes=200000))
            changed("unbounded flood log", lambda report: report["shell_cases"][0]["stream_probe"].update(log_size_bytes=400000))
            changed("lost stderr byte totals", lambda report: report["shell_cases"][0]["stream_probe"].update(actual_stderr_bytes=65536))
            changed("omitted truncation", lambda report: report["shell_cases"][0]["stream_probe"].update(stderr_truncated=False))
            changed("unproved EOF", lambda report: report["shell_cases"][0]["stream_probe"].update(streams_complete=False))
            changed("synthetic stream replacement", lambda report: report["shell_cases"][0]["stream_probe"].update(actual_run_function=False))
            for description, payload in defects:
                with self.subTest(defect=description):
                    path.write_text(json.dumps(payload), encoding="utf-8")
                    step = {"exit_code": 0, "requested_shell_paths": hosts}
                    runner.attach_child_report(step, path, "diagnostics")
                    self.assertEqual(step["exit_code"], 126)
                    self.assertIn("evidence_error", step)
                    self.assertNotIn("evidence", step)
                    failed = {"exit_code": 19, "requested_shell_paths": hosts}
                    runner.attach_child_report(failed, path, "diagnostics")
                    self.assertEqual(failed["exit_code"], 19)
            path.write_text(json.dumps(complete), encoding="utf-8")
            for requested in ([], [hosts[0]], [hosts[0], hosts[0]]):
                with self.subTest(requested=requested):
                    step = {"exit_code": 0, "requested_shell_paths": requested}
                    runner.attach_child_report(step, path, "diagnostics")
                    self.assertEqual(step["exit_code"], 126)

    def test_plan_targeted_route_runs_only_plan_without_shell_arguments(self):
        args = SimpleNamespace(layer="plan", failure_probe=None, shell_path=[Path("C:/unused/pwsh.exe")], tool_root=None)
        commands = []

        def successful_command(argv, cwd, **kwargs):
            commands.append(argv)
            return {"exit_code": 0, "stdout": "synthetic-git-identity\n", "stderr": ""}

        with patch.object(runner, "source_manifest", return_value={"synthetic-source": "hash"}), \
                patch.object(runner, "run_command", side_effect=successful_command), \
                patch.object(runner, "attach_child_report") as attach:
            report = runner.execute(args)
        self.assertEqual([step["name"] for step in report["steps"]], ["shared-plan-regression"])
        self.assertEqual([call.args[2] for call in attach.call_args_list], ["plan"])
        self.assertEqual(len(commands), 3)  # Git commit, tree, then the single targeted child.
        self.assertEqual(commands[-1][3], str(ROOT / "tests/plans/characterize_plan.py"))
        self.assertNotIn("--shell-path", commands[-1])
        self.assertTrue(report["success"])

    def test_plan_evidence_requires_all_promises_and_exact_preview_writer_identities(self):
        complete = complete_plan_report()
        with tempfile.TemporaryDirectory(prefix="wbs-plan-evidence-") as directory:
            path = Path(directory) / "plan.json"
            path.write_text(json.dumps(complete), encoding="utf-8")
            step = {"exit_code": 0}
            runner.attach_child_report(step, path, "plan")
            self.assertEqual(step["exit_code"], 0)
            self.assertEqual(step["evidence"], complete)
            self.assertIn("evidence_sha256", step)
            defects = []

            def defect(description, group, index, changes):
                report = deepcopy(complete)
                report[group][index].update(changes)
                defects.append((description, report))

            for field, value in (("schema_version", True), ("task_id", "M1-T04"), ("result", "LEVEL2_REGRESSION_PASSED"),
                                 ("success", False), ("exit_code", False), ("acceptance_ids", ["AC-027"]),
                                 ("source_unchanged", False), ("baseline_guards_preserved", False),
                                 ("input_and_neighbor_unchanged", False), ("owned_temp_removed", False),
                                 ("immutable_original_commit", "unknown"), ("import_observation", [])):
                defects.append((field, {**complete, field: value}))
            for group in ("structural_cases", "preview_cases", "parity_cases", "source_cases"):
                defects.append(("missing " + group, {**complete, group: []}))
                defects.append(("duplicate " + group, {**complete, group: complete[group][:-1] + [complete[group][0]]}))
                key = "mode" if group == "parity_cases" else "id"
                defect("wrong " + group + " ID", group, 0, {key: "unexpected"})
                defect("unhashable " + group + " ID", group, 0, {key: []})
                defect("failed " + group, group, 0, {"passed": False})
            defect("no actual structural rejection", "structural_cases", 0, {"rejected": False})
            defect("missing rejection diagnostic", "structural_cases", 0, {"error_code": ""})
            defect("whole document rejected", "structural_cases", 8, {"accepted": False})
            defect("mutable preview", "preview_cases", 0, {"mutation_rejected": False})
            defect("preview wrote files", "preview_cases", 0, {"chapter_files_written": 1})
            defect("boolean preview count", "preview_cases", 0, {"chapter_files_written": False})
            defect("aliased preview metadata", "preview_cases", 1, {"detached_from_input": False})
            defect("unbound mapping executed", "preview_cases", 1, {"unbound_execution_rejected": False})
            for group in ("parity_cases", "source_cases"):
                missing_result = deepcopy(complete)
                del missing_result[group][0]["writer_result"]
                defects.append(("missing writer result " + group, missing_result))
                defect("zero files " + group, group, 0, {"outputs": [], "written_count": 0})
                defect("missing preview " + group, group, 0, {"preview_entries": []})
                defect("false output count " + group, group, 0, {"written_count": True})
                defect("wrong coverage " + group, group, 0, {"coverage": {"complete": True, "covered_pages": 2, "section_count": 2}})
                defect("integer complete flag " + group, group, 0, {"coverage": {"complete": 1, "covered_pages": 3, "section_count": 2}})
                defect("changed neighbor " + group, group, 0, {"neighbor_unchanged": False})
                defect("replanned " + group, group, 0, {"no_replanning_or_reopening": False})
                output_changes = [{"page_ids": [3, 2]}, {"page_ids": [True]}, {"filename": "different.pdf"}, {"range": [0, 3]}]
                for index, changes in enumerate(output_changes):
                    report = deepcopy(complete)
                    report[group][0]["outputs"][0].update(changes)
                    defects.append((f"wrong physical output {group} {index}", report))
                report = deepcopy(complete)
                report[group][0]["preview_entries"][1]["start"] = 2
                defects.append(("preview gap " + group, report))
                report = deepcopy(complete)
                report[group][0]["preview_entries"][0]["sequence"] = True
                defects.append(("boolean sequence " + group, report))
            defect("executed replacement", "source_cases", 0, {"outcome": "changed_source_used"})
            defect("replacement changed", "source_cases", 0, {"replacement_unchanged": False})
            defect("wrong bound digest", "source_cases", 0, {"captured_sha256": "c" * 64})
            defect("unchanged replacement", "source_cases", 0, {"replacement_sha256": "a" * 64})
            defect("invalid replacement digest", "source_cases", 0, {"replacement_sha256": "not-a-digest"})
            defect("nonhex replacement digest", "source_cases", 0, {"replacement_sha256": "g" * 64})
            defect("replacement page content used", "source_cases", 0, {"page_content_sha256": ["f" * 64] * 3})
            defect("missing original page content", "source_cases", 0, {"original_page_content_sha256": []})
            defect("ordinary writer page content mismatch", "parity_cases", 0, {"page_content_sha256": ["f" * 64] * 3})
            defect("ordinary original content count mismatch", "parity_cases", 0, {"original_page_content_sha256": []})
            defect("ordinary original content invalid hash", "parity_cases", 0, {"original_page_content_sha256": ["not-a-hash"] * 3})
            for group in ("parity_cases", "source_cases"):
                for field in ("original_page_content_sha256", "page_content_sha256"):
                    missing_content = deepcopy(complete)
                    del missing_content[group][0][field]
                    defects.append(("omitted " + group + " " + field, missing_content))
            defect("uncaptured source", "source_cases", 0, {"source_identity": {"binding": "path"}})
            for group in ("parity_cases", "source_cases"):
                report = deepcopy(complete)
                case = report[group][0]
                case["writer_result"] = {"mode": case.get("mode", "manual"), "total_pages": 3,
                                         "written_count": 2, "coverage": deepcopy(case["coverage"]),
                                         "source_identity": deepcopy(case["source_identity"]),
                                         "outputs": [{**entry, "page_count": entry["end"] - entry["start"]}
                                                     for entry in case["preview_entries"]]}
                case["writer_result"] = transaction_result(case["writer_result"], 1 if group == "parity_cases" else 4)
                path.write_text(json.dumps(report), encoding="utf-8")
                valid_result_step = {"exit_code": 0}
                runner.attach_child_report(valid_result_step, path, "plan")
                self.assertEqual(valid_result_step["exit_code"], 0)
                for field in ("mode", "total_pages", "written_count", "coverage", "source_identity", "outputs"):
                    changed = deepcopy(report)
                    del changed[group][0]["writer_result"][field]
                    defects.append(("omitted writer result " + group + " " + field, changed))
                for field, value in (("written_count", 1), ("total_pages", True), ("outputs", []),
                                     ("coverage", {}), ("source_identity", {}), ("mode", "invalid"), ("mode", [])):
                    changed = deepcopy(report)
                    changed[group][0]["writer_result"][field] = value
                    defects.append(("writer result " + group + " " + field, changed))
                changed = deepcopy(report)
                changed[group][0]["writer_result"]["outputs"][0]["page_count"] = True
                defects.append(("boolean result page count " + group, changed))
            missing_deleted = deepcopy(complete)
            del missing_deleted["source_cases"][1]["deleted_path_snapshot"]
            defects.append(("omitted deleted-path execution", missing_deleted))
            defect("invalid deleted-path object", "source_cases", 1, {"deleted_path_snapshot": []})
            for field in ("writer_result", "original_page_content_sha256", "page_content_sha256", "source_identity"):
                missing_nested = deepcopy(complete)
                del missing_nested["source_cases"][1]["deleted_path_snapshot"][field]
                defects.append(("omitted deleted-path " + field, missing_nested))
            for field, value in (("written_count", 0), ("coverage", {}), ("outputs", []),
                                 ("page_content_sha256", ["f" * 64] * 3)):
                changed = deepcopy(complete)
                changed["source_cases"][1]["deleted_path_snapshot"][field] = value
                defects.append(("deleted-path mismatch " + field, changed))
            changed = deepcopy(complete)
            nested = changed["source_cases"][1]["deleted_path_snapshot"]
            nested["source_identity"]["sha256"] = "f" * 64
            nested["writer_result"]["source_identity"]["sha256"] = "f" * 64
            defects.append(("deleted-path switched bound source", changed))
            changed = deepcopy(complete)
            nested = changed["source_cases"][1]["deleted_path_snapshot"]
            nested["original_page_content_sha256"] = nested["page_content_sha256"] = ["f" * 64] * 3
            defects.append(("deleted-path switched original content", changed))
            changed = deepcopy(complete)
            changed["source_cases"][1]["deleted_path_snapshot"]["writer_result"]["outputs"] = []
            defects.append(("deleted-path writer result mismatch", changed))
            for changes in ({"seed": 0}, {"count": 299}, {"count": True}, {"passed": False},
                            {"per_mode": {"manual": 100, "1": 100}}, {"per_mode": {"manual": 99, "1": 101, "2": 100}}):
                defects.append(("seeded modes " + repr(changes), {**complete, "seeded_plan_cases": {**complete["seeded_plan_cases"], **changes}}))
            for description, payload in defects:
                with self.subTest(defect=description):
                    path.write_text(json.dumps(payload), encoding="utf-8")
                    step = {"exit_code": 0}
                    runner.attach_child_report(step, path, "plan")
                    self.assertEqual(step["exit_code"], 126)
                    self.assertIn("evidence_error", step)
                    self.assertNotIn("evidence", step)
                    failed_child = {"exit_code": 19}
                    runner.attach_child_report(failed_child, path, "plan")
                    self.assertEqual(failed_child["exit_code"], 19)

    def test_bookmark_evidence_requires_complete_targets_normalization_and_preservation(self):
        complete = {
            "schema_version": 1, "task_id": "M1-T03", "result": "LEVEL1_REGRESSION_PASSED",
            "success": True, "exit_code": 0, "acceptance_ids": ["AC-019", "AC-020", "AC-021", "AC-022"],
            "engine_case_count": 4,
            "engine_cases": [{"oracle_id": name, "passed": True} for name in ("BM-01", "BM-02", "BM-05", "BM-08")],
            "normalization_cases": [{"id": name, "passed": True} for name in (
                "invalid-destinations", "deep-lineage", "cyclic-outline", "malformed-outline", "traversal-limits")],
            "seeded_level1_cases": {"seed": 20261009, "count": 150, "passed": True},
            "sampled_writer_cases": [{"index": number, "outputs": [{"filename": "synthetic.pdf"}]}
                                     for number in range(10)],
            "source_unchanged": True, "baseline_guards_preserved": True,
            "input_and_neighbor_unchanged": True, "owned_temp_removed": True,
            "immutable_original_commit": "0de84f367f9bd5ddfa3f408a9c29505d7a39633f",
            "import_observation": {"import_safe": True},
        }
        with tempfile.TemporaryDirectory(prefix="wbs-bookmark-evidence-") as directory:
            path = Path(directory) / "bookmarks.json"
            path.write_text(json.dumps(complete), encoding="utf-8")
            step = {"exit_code": 0}
            runner.attach_child_report(step, path, "bookmarks")
            self.assertEqual(step["exit_code"], 0)
            self.assertEqual(step["evidence"], complete)
            self.assertIn("evidence_sha256", step)
            defects = {
                "boolean schema": {"schema_version": True},
                "wrong task": {"task_id": "M1-T02"},
                "wrong result": {"result": "MANUAL_REGRESSION_PASSED"},
                "claimed failure": {"success": False},
                "boolean exit": {"exit_code": False},
                "wrong case count": {"engine_case_count": 3},
                "count without records": {"engine_cases": []},
                "duplicate oracle IDs": {"engine_cases": complete["engine_cases"][:-1] + [complete["engine_cases"][0]]},
                "unknown oracle ID": {"engine_cases": complete["engine_cases"][:-1] + [{"oracle_id": "BM-03", "passed": True}]},
                "invalid oracle ID type": {"engine_cases": complete["engine_cases"][:-1] + [{"oracle_id": [], "passed": True}]},
                "failed target": {"engine_cases": complete["engine_cases"][:-1] + [{"oracle_id": "BM-08", "passed": False}]},
                "missing normalization": {"normalization_cases": []},
                "duplicate normalization ID": {"normalization_cases": complete["normalization_cases"][:-1]
                                               + [complete["normalization_cases"][0]]},
                "unknown normalization ID": {"normalization_cases": complete["normalization_cases"][:-1]
                                             + [{"id": "unexpected", "passed": True}]},
                "invalid normalization ID type": {"normalization_cases": complete["normalization_cases"][:-1]
                                                  + [{"id": [], "passed": True}]},
                "failed normalization": {"normalization_cases": complete["normalization_cases"][:-1]
                                         + [{**complete["normalization_cases"][-1], "passed": False}]},
                "wrong seed": {"seeded_level1_cases": {"seed": 0, "count": 150, "passed": True}},
                "wrong property count": {"seeded_level1_cases": {"seed": 20261009, "count": 0, "passed": True}},
                "failed properties": {"seeded_level1_cases": {"seed": 20261009, "count": 150, "passed": False}},
                "missing writer samples": {"sampled_writer_cases": []},
                "duplicate writer indexes": {"sampled_writer_cases": complete["sampled_writer_cases"][:-1]
                                             + [complete["sampled_writer_cases"][0]]},
                "boolean writer index": {"sampled_writer_cases": complete["sampled_writer_cases"][:-1]
                                         + [{"index": True, "outputs": [{"filename": "synthetic.pdf"}]}]},
                "zero output files": {"sampled_writer_cases": complete["sampled_writer_cases"][:-1]
                                      + [{"index": 9, "outputs": []}]},
                "incomplete acceptance": {"acceptance_ids": ["AC-019"]},
                "changed source": {"source_unchanged": False},
                "changed guards": {"baseline_guards_preserved": False},
                "changed inputs or neighbors": {"input_and_neighbor_unchanged": False},
                "missing cleanup": {"owned_temp_removed": False},
                "wrong immutable reference": {"immutable_original_commit": "unknown"},
                "unsafe import": {"import_observation": {"import_safe": False}},
                "invalid import observation": {"import_observation": []},
            }
            for description, changes in defects.items():
                with self.subTest(defect=description):
                    path.write_text(json.dumps({**complete, **changes}), encoding="utf-8")
                    step = {"exit_code": 0}
                    runner.attach_child_report(step, path, "bookmarks")
                    self.assertEqual(step["exit_code"], 126)
                    self.assertIn("evidence_error", step)
                    self.assertNotIn("evidence", step)
                    step = {"exit_code": 19}
                    runner.attach_child_report(step, path, "bookmarks")
                    self.assertEqual(step["exit_code"], 19)

    def test_level2_evidence_requires_complete_targets_hierarchy_and_preservation(self):
        complete = {
            "schema_version": 1, "task_id": "M1-T04", "result": "LEVEL2_REGRESSION_PASSED",
            "success": True, "exit_code": 0, "acceptance_ids": ["AC-023", "AC-024", "AC-025", "AC-026"],
            "engine_case_count": 4,
            "engine_cases": [{"oracle_id": name, "passed": True} for name in ("BM-03", "BM-04", "BM-06", "BM-07")],
            "hierarchy_cases": [{"id": name, "passed": True} for name in (
                "duplicate-parent-subtrees", "invalid-parent-subtrees", "child-order-and-aliases",
                "invalid-child-destinations", "deep-and-malformed-outlines", "no-usable-level2")],
            "seeded_level2_cases": {"seed": 20261009, "count": 150, "passed": True},
            "sampled_writer_cases": [{"index": number, "outputs": [{"filename": "synthetic.pdf"}]}
                                     for number in range(10)],
            "source_unchanged": True, "baseline_guards_preserved": True,
            "input_and_neighbor_unchanged": True, "owned_temp_removed": True,
            "immutable_original_commit": "0de84f367f9bd5ddfa3f408a9c29505d7a39633f",
            "import_observation": {"import_safe": True},
        }
        with tempfile.TemporaryDirectory(prefix="wbs-level2-evidence-") as directory:
            path = Path(directory) / "level2.json"
            path.write_text(json.dumps(complete), encoding="utf-8")
            step = {"exit_code": 0}
            runner.attach_child_report(step, path, "level2")
            self.assertEqual(step["exit_code"], 0)
            self.assertEqual(step["evidence"], complete)
            self.assertIn("evidence_sha256", step)
            defects = {
                "boolean schema": {"schema_version": True},
                "wrong task": {"task_id": "M1-T03"},
                "wrong result": {"result": "LEVEL1_REGRESSION_PASSED"},
                "claimed failure": {"success": False},
                "boolean exit": {"exit_code": False},
                "incomplete acceptance": {"acceptance_ids": ["AC-023"]},
                "wrong case count": {"engine_case_count": 3},
                "count without actual cases": {"engine_cases": []},
                "duplicate target ID": {"engine_cases": complete["engine_cases"][:-1] + [complete["engine_cases"][0]]},
                "wrong target ID": {"engine_cases": complete["engine_cases"][:-1] + [{"oracle_id": "BM-08", "passed": True}]},
                "invalid target ID type": {"engine_cases": complete["engine_cases"][:-1] + [{"oracle_id": [], "passed": True}]},
                "failed target": {"engine_cases": complete["engine_cases"][:-1] + [{"oracle_id": "BM-07", "passed": False}]},
                "missing hierarchy cases": {"hierarchy_cases": []},
                "duplicate hierarchy ID": {"hierarchy_cases": complete["hierarchy_cases"][:-1] + [complete["hierarchy_cases"][0]]},
                "wrong hierarchy ID": {"hierarchy_cases": complete["hierarchy_cases"][:-1] + [{"id": "unexpected", "passed": True}]},
                "invalid hierarchy ID type": {"hierarchy_cases": complete["hierarchy_cases"][:-1] + [{"id": [], "passed": True}]},
                "failed hierarchy case": {"hierarchy_cases": complete["hierarchy_cases"][:-1]
                                          + [{**complete["hierarchy_cases"][-1], "passed": False}]},
                "wrong seed": {"seeded_level2_cases": {"seed": 0, "count": 150, "passed": True}},
                "wrong property count": {"seeded_level2_cases": {"seed": 20261009, "count": 0, "passed": True}},
                "failed properties": {"seeded_level2_cases": {"seed": 20261009, "count": 150, "passed": False}},
                "missing writer samples": {"sampled_writer_cases": []},
                "duplicate writer indexes": {"sampled_writer_cases": complete["sampled_writer_cases"][:-1]
                                             + [complete["sampled_writer_cases"][0]]},
                "boolean writer index": {"sampled_writer_cases": complete["sampled_writer_cases"][:-1]
                                         + [{"index": True, "outputs": [{"filename": "synthetic.pdf"}]}]},
                "zero writer files": {"sampled_writer_cases": complete["sampled_writer_cases"][:-1]
                                      + [{"index": 9, "outputs": []}]},
                "changed source": {"source_unchanged": False},
                "changed guards": {"baseline_guards_preserved": False},
                "changed inputs or neighbors": {"input_and_neighbor_unchanged": False},
                "missing cleanup": {"owned_temp_removed": False},
                "wrong immutable reference": {"immutable_original_commit": "unknown"},
                "unsafe import": {"import_observation": {"import_safe": False}},
                "invalid import observation": {"import_observation": []},
            }
            for description, changes in defects.items():
                with self.subTest(defect=description):
                    path.write_text(json.dumps({**complete, **changes}), encoding="utf-8")
                    step = {"exit_code": 0}
                    runner.attach_child_report(step, path, "level2")
                    self.assertEqual(step["exit_code"], 126)
                    self.assertIn("evidence_error", step)
                    self.assertNotIn("evidence", step)
                    step = {"exit_code": 19}
                    runner.attach_child_report(step, path, "level2")
                    self.assertEqual(step["exit_code"], 19)

    def test_existing_report_is_preserved(self):
        with tempfile.TemporaryDirectory(prefix="wbs-no-overwrite-") as directory:
            path = Path(directory) / "existing.json"
            path.write_bytes(b"preexisting evidence")
            result = runner.run_command([sys.executable, "-I", "-B", str(SCRIPT),
                                         "--layer", "python", "--failure-probe", "python",
                                         "--report", str(path)], Path(directory))
            self.assertEqual(result["exit_code"], 1)
            self.assertEqual(path.read_bytes(), b"preexisting evidence")

    def test_rejected_child_receipt_retains_raw_failure_without_acceptance_evidence(self):
        with tempfile.TemporaryDirectory(prefix="wbs-rejected-receipt-") as directory:
            path = Path(directory) / "partial.json"
            partial = {"schema_version": 1, "success": False, "cleanup_safe": False,
                       "cases": [], "failure": "owned tree stop unproved"}
            path.write_text(json.dumps(partial), encoding="utf-8")
            step = {"exit_code": 0}
            runner.attach_child_report(step, path, "outcomes")
            self.assertEqual(step["exit_code"], 126)
            self.assertIn("evidence_error", step)
            self.assertNotIn("evidence", step)
            self.assertEqual(step["failed_evidence"], partial)
            self.assertEqual(step["failed_evidence_sha256"], runner.sha256(path.read_bytes()))
            self.assertIs(step["cleanup_safe"], False)

    def test_reports_require_absolute_external_paths(self):
        with self.assertRaises(ValueError):
            runner.new_external_path(Path("relative.json"))
        with self.assertRaises(ValueError):
            runner.new_external_path(ROOT / "must-not-create.json")
        self.assertFalse((ROOT / "must-not-create.json").exists())

    def test_unverifiable_ebook_report_never_authorizes_outer_cleanup(self):
        with tempfile.TemporaryDirectory(prefix="wbs-ebook-receipt-") as directory:
            path = Path(directory) / "child.json"
            for data in (None, b"{bad-json", b'{"schema_version":1,"success":true}'):
                with self.subTest(receipt=data):
                    if data is not None:
                        path.write_bytes(data)
                    step = {"exit_code": 0}
                    runner.attach_child_report(step, path, "ebooks")
                    self.assertEqual(step["exit_code"], 126)
                    self.assertIs(step["cleanup_safe"], False)
                    self.assertNotIn("evidence", step)
                    self.assertIn("evidence_error", step)
            with patch.object(runner.tempfile, "TemporaryDirectory") as temporary:
                temporary.return_value.name = directory
                with runner.owned_evidence_workspace([step]):
                    pass
                temporary.return_value._finalizer.detach.assert_called_once()
                temporary.return_value.cleanup.assert_not_called()

    def test_shell_bootstrap_passes_paths_as_data(self):
        with tempfile.TemporaryDirectory(prefix="wbs-shell-data-") as directory:
            work = Path(directory) / "quote' [space] å"
            work.mkdir()
            with patch.dict(os.environ, {"PSModulePath": "untrusted foreign modules"}):
                command, environment = runner.shell_command(Path("C:/trusted/pwsh.exe"),
                                                           work, work / "evidence.json", work, False)
            self.assertEqual(environment["WBS_TEST_WORK"], str(work))
            self.assertEqual(environment["PSModulePath"], str(Path("C:/trusted/Modules")))
            self.assertNotIn(str(work), " ".join(command))
            self.assertIn("-EncodedCommand", command)
            self.assertEqual(command[command.index("-ExecutionPolicy") + 1], "RemoteSigned")
            self.assertNotIn("Bypass", command)

    def test_git_provenance_drops_inherited_repository_selectors(self):
        with patch.dict(os.environ, {name: "hostile-unrelated-value" for name in runner.GIT_SELECTORS}):
            environment = runner.child_environment()
            self.assertFalse(any(name in environment for name in runner.GIT_SELECTORS))
            self.assertTrue(all(name in os.environ for name in runner.GIT_SELECTORS))


if __name__ == "__main__":
    unittest.main(verbosity=2)
