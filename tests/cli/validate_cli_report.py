"""Independent strict actual-host CLI receipt contract; IDs alone never pass."""

from hashlib import sha256
import json
import math
from pathlib import Path
import re
import importlib.util

spec = importlib.util.spec_from_file_location("wbs_cli_console_validation", Path(__file__).resolve().parents[2] / "tests/manual/current_launchers.py")
console = importlib.util.module_from_spec(spec)
spec.loader.exec_module(console)

ACCEPTANCE = ["AC-053", "AC-054", "AC-055", "AC-056"]
EXECUTIONS = {"execute-manual", "execute-auto1", "execute-auto2", "execute-manual-nopause", "execute-epub-keep"}
PREVIEWS = {"preview-manual", "preview-auto1", "preview-auto2", "preview-auto1-without-ni", "preview-epub", "preview-azw3"}
NO_PLANS = {"no-plan-auto1", "no-plan-auto2", "no-plan-parents"}
EARLY_ARGUMENTS = {"missing-input", "missing-mode", "missing-auto-level", "missing-manual-starts",
    "auto-with-starts", "manual-with-level", "level-without-mode", "starts-without-mode", "bad-mode", "bad-level",
    "bad-manual-token", "empty-manual-token", "version-with-processing", "preview-with-retention",
    "preview-missing-mode", "explicit-manual-missing-starts", "unknown-argument"}
DEPENDENCIES = {"missing-python", "missing-calibre"}
KINDS = [*sorted(EXECUTIONS), *sorted(PREVIEWS), *sorted(NO_PLANS), *sorted(EARLY_ARGUMENTS), *sorted(DEPENDENCIES),
         "manual-out-of-range", "version", "help"]
EXPECTED = {kind: 0 if kind in EXECUTIONS | PREVIEWS | {"version", "help"} else 5 if kind in NO_PLANS
            else 3 if kind in DEPENDENCIES else 2 for kind in KINDS}
RANGES = {"manual": [[0, 2], [2, 4], [4, 6]], "1": [[0, 2], [2, 4], [4, 6]], "2": [[0, 1], [1, 2], [2, 3], [3, 4], [4, 5], [5, 6]]}
PARAMETERS = {"InputFile", "OutputDirectory", "Mode", "BookmarkLevel", "StartPages", "Preview", "NonInteractive",
              "NoPause", "Version", "PythonPath", "CalibrePath", "KeepConvertedPdf"}
APPLICATION = {"WinBookSplit.ps1", "WinBookSplit.bat", "requirements.txt", "engine/WinBookSplit.Paths.ps1",
    "engine/WinBookSplit.Diagnostics.ps1", "engine/WinBookSplit.Runtime.ps1", "engine/WinBookSplit.Process.ps1",
    "engine/WinBookSplit.Outcomes.json", "engine/winbooksplit_engine.py", "engine/winbooksplit_windows.py",
    "engine/winbooksplit_conversion.py", "engine/winbooksplit_job.py", "engine/WinBookSplit.Logging.ps1",
    "engine/WinBookSplit.Support.ps1", "Export-WinBookSplitDiagnostics.ps1"}
HEX = re.compile(r"[0-9a-f]{64}")


def need(condition, message):
    if not condition:
        raise ValueError(message)


def hashes(value):
    return isinstance(value, dict) and bool(value) and all(isinstance(name, str) and isinstance(digest, str)
        and HEX.fullmatch(digest) for name, digest in value.items())


def exact_cases(rows, expected):
    need(isinstance(rows, list) and len(rows) == len(expected) and all(isinstance(row, dict) for row in rows), "CLI case count missing")
    ids = [row.get("id") for row in rows]
    need(len(ids) == len(set(ids)) and set(ids) == set(expected), "CLI cases missing, duplicate or unexpected")
    return {row["id"]: row for row in rows}


def identities(observed):
    need(isinstance(observed, dict) and set(observed) == {"source", "neighbor", "prior"}, "Source/neighbor/prior identity observations missing")
    for item in observed.values():
        need(isinstance(item, dict) and isinstance(item.get("sha256"), str) and HEX.fullmatch(item["sha256"])
            and all(type(item.get(field)) is int and item[field] >= 0 for field in ("size_bytes", "device", "inode", "attributes"))
            and item["size_bytes"] > 0 and item["inode"] > 0, "Concrete source identity/hash/attributes missing")


def coverage(plan, expected_ranges):
    need(isinstance(plan, dict) and type(plan.get("total_pages")) is int and plan["total_pages"] > 0,
         "Plan physical page count missing")
    entries = plan.get("entries")
    need(isinstance(entries, list) and len(entries) == len(expected_ranges) and entries,
         "Plan section count missing")
    previous, names = 0, []
    for number, entry in enumerate(entries, 1):
        need(isinstance(entry, dict) and type(entry.get("start")) is int and type(entry.get("end")) is int
            and entry["start"] == previous < entry["end"] <= plan["total_pages"]
            and entry.get("sequence") == number and isinstance(entry.get("title"), str)
            and isinstance(entry.get("filename"), str) and entry["filename"].endswith(".pdf"),
            "Plan has gaps, overlaps, empty ranges or incomplete section data")
        previous = entry["end"]
        names.append(entry["filename"].casefold())
    need(previous == plan["total_pages"] and len(names) == len(set(names))
        and [[entry["start"], entry["end"]] for entry in entries] == expected_ranges
        and plan.get("ranges") == expected_ranges
        and plan.get("coverage") == {"complete": True, "covered_pages": plan["total_pages"], "section_count": len(entries)},
        "Plan lacks exact ordered complete physical coverage")


def validate_case(case):
    need(isinstance(case, dict) and case.get("kind") in EXPECTED, "Unknown CLI case")
    kind = case["kind"]
    version = case.get("host_version", "")
    prefix = "PS51" if version.startswith("5.1.") else "PS7" if version.startswith("7.") else None
    need(prefix and case.get("id") == prefix + "-" + kind, "Actual CLI host/version ID mismatch")
    need(all(case.get(field) is True for field in ("passed", "actual_process", "source_read_only_observed",
        "source_neighbor_prior_unchanged", "application_unchanged", "known_output_cleanup_complete")), "Actual CLI preservation/cleanup proofs absent")
    need(type(case.get("exit_code")) is int and case["exit_code"] == EXPECTED[kind] and case.get("timed_out") is False
        and case.get("stdin") == ("PIPE" if kind == "execute-manual-nopause" else "DEVNULL")
        and (kind != "execute-manual-nopause" or case.get("stdin_utf8") == "Y\nN\n")
        and type(case.get("pid")) is int and case["pid"] > 0
        and type(case.get("elapsed_seconds")) in (int, float) and math.isfinite(case["elapsed_seconds"])
        and 0 <= case["elapsed_seconds"] < case.get("timeout_seconds", 0), "Actual native CLI exit/stdin/deadline proof missing")
    command, parameters = case.get("command"), case.get("parameters")
    need(isinstance(command, list) and command and command[0] == case.get("shell_executable")
        and "-NoProfile" in command and "-File" in command
        and (("-NonInteractive" in command[:command.index("-File")]) == (kind != "execute-manual-nopause"))
        and isinstance(parameters, list) and all(isinstance(value, str) for value in command + parameters), "Actual literal CLI command missing")
    if kind != "help":
        need(command[command.index("-File") + 2:] == parameters and Path(command[command.index("-File") + 1]).name == "WinBookSplit.ps1", "Native CLI arguments differ from receipt")
    for field in ("stdout", "stderr"):
        need(isinstance(case.get(field), str) and case.get(field + "_sha256") == sha256(case[field].encode("utf-8")).hexdigest(), "Native CLI stream/hash mismatch")
    text = case["stdout"] + case["stderr"]
    forbidden = ("Read-Host", "Read and Prompt functionality", "Press Enter to exit", "[FALLBACK]", "\x1b[2J")
    if kind != "execute-manual-nopause":
        forbidden += ("Open output folder", "Create these chapter PDFs?", "Open the generated PDF?")
    else:
        need("Create these chapter PDFs?" in text and "Open output folder?" in text
             and "No chapter PDFs have been written." in text, "NoPause bypassed interactive confirmation")
    need(not any(fragment in text for fragment in forbidden), "CLI attempted a forbidden prompt, pause, fallback, folder launch or clear")
    for key in ("source_observations_before", "source_observations_after", "source_observations_after_cleanup"):
        identities(case.get(key))
    need(case["source_observations_before"] == case["source_observations_after"] == case["source_observations_after_cleanup"]
        and case["source_observations_before"]["source"]["attributes"] & 1, "Source/neighbor/prior changed or read-only control absent")
    need(case.get("output_members_before") == case.get("output_members_after_cleanup") == ["prior-output.pdf"], "Cleanup reached prior output or retained unexplained members")
    app = case.get("application_sha256")
    expected_app = {"WinBookSplit.ps1", "engine/WinBookSplit.Outcomes.json"} if kind == "version" else APPLICATION
    need(hashes(app) and set(app) == expected_app and case.get("application_members") == sorted(app), "Copied shipped application scope/hash missing")
    if kind in {"version", "help"}:
        need(case.get("outcome") is None and case.get("engine_records") == [] and case.get("process_summaries") == []
            and case.get("log_sha256") is None and case.get("console_evidence") is None and case.get("output_members_after") == ["prior-output.pdf"], "Version/help performed processing or wrote output")
        if kind == "version":
            need(parameters == ["-Version"] and case["stdout"].strip() == "WinBookSplit 1.0.0-dev" and not case["stderr"], "Dependency-independent canonical version missing")
        else:
            help_receipt = case.get("help")
            need(isinstance(help_receipt, dict) and help_receipt.get("host_version") == version
                and PARAMETERS <= set(help_receipt.get("parameters", []))
                and all(fragment in help_receipt.get("text", "") for fragment in ("-StartPages", "-NonInteractive", "-Preview", "-Version", "physical")), "Actual full comment-based CLI help incomplete")
        return
    final = case.get("outcome")
    need(isinstance(final, dict) and final.get("protocol") == "winbooksplit.outcome" and type(final.get("version")) is int
        and final["version"] == 1 and type(final.get("exit_code")) is int and final["exit_code"] == EXPECTED[kind]
        and type(final.get("written_count")) is int, "Authoritative CLI outcome incomplete")
    printed = [line[len("[OUTCOME] "):] for line in case["stdout"].splitlines() if line.startswith("[OUTCOME] ")]
    need(len(printed) == 1 and json.loads(printed[0]) == final, "Actual stdout/outcome differs or duplicates")
    frame = final.get("engine_result")
    if kind in EARLY_ARGUMENTS | DEPENDENCIES:
        need(frame is None and case["engine_records"] == [] and case["process_summaries"] == []
            and case["log_sha256"] is None and case["output_members_after"] == ["prior-output.pdf"], "Early refusal reached engine, log or writes")
        need(case.get("console_evidence") is None, "Early refusal invented a run record")
        need(final["written_count"] == 0 and final.get("final_directory") is None, "Early refusal fabricated publication")
        need("[DEPENDENCY]" not in case["stdout"] and "[ENGINE]" not in case["stdout"], "Early refusal selected processing")
        if kind in EARLY_ARGUMENTS:
            need(final.get("code") == "invalid_arguments" and "[DEPENDENCY-ERROR]" not in case["stdout"], "Argument contradiction did not reject before discovery")
        else:
            need(final.get("code") == ("runtime_invalid" if kind == "missing-python" else "converter_invalid")
                and "[DEPENDENCY-ERROR]" in case["stdout"], "Missing explicit dependency was not actionable dependency3")
        return
    need(isinstance(frame, dict) and frame.get("protocol") == "winbooksplit.result" and frame.get("version") == 1
        and frame.get("exit_code") == EXPECTED[kind] and frame.get("mode") == final.get("mode")
        and frame.get("status") == final.get("status") and frame.get("code") == final.get("code"), "Final CLI/native/engine attempt binding differs")
    if kind in PREVIEWS:
        plan = case.get("plan")
        need(frame.get("status") == "preview" and frame.get("code") == "preview_complete"
            and frame.get("written_count") == final["written_count"] == 0 and frame.get("execution") is None
            and final.get("final_directory") is None and frame.get("fallback_modes") == [] and frame.get("plan") == plan,
            "Preview reports execution, writes or fallback")
        coverage(plan, case.get("expected_ranges"))
        need(plan["entries"] == case.get("expected_entries") and plan.get("output_naming", {}).get("resolved_base") == case.get("output_base")
            and case["output_members_after"] == ["prior-output.pdf"] and case["log_sha256"] is None,
            "Preview differs from bound filenames/ranges or wrote files")
        need(case.get("console_evidence") is None, "Preview invented a persisted run record")
        need("Preview only:" in case["stdout"] and "Done." not in case["stdout"] and "Output: " not in case["stdout"], "Plan-only disclosure missing or false completion")
        if kind in {"preview-epub", "preview-azw3"}:
            generated = plan.get("conversion", {}).get("generated_pdf_identity", {})
            reference = case.get("conversion_reference")
            need(case.get("temporary_conversion_absent") is True and isinstance(reference, dict)
                and reference.get("process", {}).get("exit_code") == 0
                and generated.get("page_content_sha256") == reference.get("page_content_sha256") == case.get("expected_content_sha256")
                and generated.get("size_bytes", 0) > 0 and isinstance(generated.get("sha256"), str) and HEX.fullmatch(generated["sha256"])
                and "temporary" in case["stdout"].lower() and "clean" in case["stdout"].lower(), "Real ebook preview conversion/content/cleanup disclosure missing")
            need(plan.get("original_ebook_identity", {}).get("sha256") == case["source_observations_before"]["source"]["sha256"], "Ebook preview forgot original immutable source")
        return
    transports = case.get("process_summaries")
    console.validate_console_evidence(case.get("console_evidence"), final, case.get("log_sha256"),
        source_path=case.get("input_path"), source_observation=case["source_observations_before"]["source"])
    need(case.get("engine_records") == [frame] and isinstance(transports, list) and len(transports) == 1
        and all(transports[0].get(field) is True for field in ("JobAssigned", "ParentStopped", "DescendantsStopped", "StreamsComplete"))
        and transports[0].get("ExitCode") == EXPECTED[kind] and transports[0].get("ResultRecordCount") == 1
        and isinstance(case.get("log_sha256"), str) and HEX.fullmatch(case["log_sha256"]), "Actual logged engine/tree/EOF receipt missing")
    if kind in EXECUTIONS:
        execution = frame.get("execution")
        pages = 3 if kind == "execute-epub-keep" else 6
        ranges = [[0, 1], [1, 2], [2, 3]] if kind == "execute-epub-keep" else RANGES[frame["mode"]]
        need(frame.get("status") == "success" and frame.get("code") == "split_complete" and isinstance(execution, dict)
            and final["written_count"] == frame.get("written_count") == execution.get("written_count") == len(ranges)
            and final.get("final_directory") == execution.get("final_directory")
            and Path(execution["final_directory"]).parent == Path(case["output_base"]), "CLI success lacks positive exact publication")
        publication, entries = case.get("final_publication"), case.get("expected_entries")
        need(isinstance(publication, list) and case.get("expected_ranges") == ranges and [item.get("range") for item in publication] == ranges
            and [page for item in publication for page in item.get("page_ids", [])] == list(range(1, pages + 1))
            and [item.get("filename") for item in publication] == [entry.get("filename") for entry in entries]
            and all(item.get("page_count") == item["range"][1] - item["range"][0]
                and item.get("size_bytes", 0) > 0 and HEX.fullmatch(item.get("sha256", "")) for item in publication), "Native CLI physical pages/files differ from plan")
        need(case.get("publication_manifest") == execution.get("manifest") and execution["manifest"].get("outputs") == execution.get("outputs")
            and isinstance(case.get("publication_manifest_sha256"), str) and HEX.fullmatch(case["publication_manifest_sha256"])
            and case.get("expected_content_sha256") == case.get("output_content_sha256") and len(case["output_content_sha256"]) == pages
            and "Done." in case["stdout"] and "Output: " in case["stdout"], "Native CLI manifest/content/completion missing")
        if kind == "execute-epub-keep":
            retained = execution.get("retained_intermediate")
            need(isinstance(retained, dict) and retained.get("filename") == "WinBookSplit_Converted.pdf"
                and retained.get("page_count") == pages and retained.get("size_bytes", 0) > 0 and HEX.fullmatch(retained.get("sha256", ""))
                and case.get("retained_sha256_verified") is True and case.get("retained_page_content_verified") is True
                and execution.get("conversion", {}).get("generated_pdf_identity", {}).get("page_content_sha256") == case["expected_content_sha256"]
                and execution.get("original_ebook_identity", {}).get("sha256") == case["source_observations_before"]["source"]["sha256"], "CLI retained ebook example lacks exact full PDF/content/source verification")
    else:
        need(frame.get("written_count") == final["written_count"] == 0 and frame.get("execution") is None
            and final.get("final_directory") is None and case.get("final_publication") is None
            and "Done." not in case["stdout"] and "Output: " not in case["stdout"], "No-plan/invalid manual falsely succeeded")
        need(frame.get("status") == ("no_plan" if kind in NO_PLANS else "invalid_input"), "Requested plan failure status differs")
        need(frame.get("code") == ("no_bookmarks_at_level" if kind == "no-plan-parents" else "no_bookmarks" if kind in NO_PLANS else "invalid_start_pages"), "Requested mode failure code differs")


def validate_cli_report(report, requested_shells, calibre):
    need(isinstance(report, dict) and report.get("schema_version") == 1 and report.get("task_id") == "M3-T02"
        and report.get("result") == "CLI_REGRESSION_PASSED" and report.get("success") is True and report.get("exit_code") == 0,
        "CLI report did not pass")
    need(report.get("acceptance_ids") == ACCEPTANCE and all(report.get(field) is True for field in
        ("source_unchanged", "baseline_guards_preserved", "machine_settings_unchanged", "owned_temp_removed")), "CLI top-level acceptance/cleanup evidence missing")
    manifest = report.get("tested_path_sha256")
    need(hashes(manifest) and APPLICATION | {"tests/cli/characterize_cli.py", "tests/cli/validate_cli_report.py", "tests/cli/Probe-Help.ps1", "tests/run_tests.py"} <= set(manifest), "Promised CLI/source/runner raw hashes missing")
    hosts = exact_cases(report.get("host_cases"), {"PS51", "PS7"})
    need(len(requested_shells) == 2 and len(set(requested_shells)) == 2
        and {host.get("shell_executable") for host in hosts.values()} == set(requested_shells), "Actual requested CLI host paths differ")
    for name, host in hosts.items():
        need(host.get("passed") is True and host.get("exit_code") == 0 and host.get("host_major") == (5 if name == "PS51" else 7)
            and host.get("syntax_error_count") == 0 and host.get("stored_policies") == host.get("policies_after")
            and isinstance(host.get("stored_policies"), list) and bool(host["stored_policies"]), "Actual host or unchanged stored-policy proof missing")
    cases = exact_cases(report.get("cases"), {name + "-" + kind for name in hosts for kind in KINDS})
    need(report.get("case_count") == len(cases), "CLI count contradicts exact controls")
    for row in cases.values():
        validate_case(row)
        host = hosts[row["id"].split("-", 1)[0]]
        need(row["host_version"] == host["host_version"] and row["shell_executable"] == host["shell_executable"], "CLI case uses unrequested host")
        need(all(manifest[name] == digest for name, digest in row["application_sha256"].items()), "Actual copied CLI bytes differ from tested source manifest")
    fixtures = report.get("fixture_provenance")
    need(isinstance(fixtures, dict) and fixtures.get("authored_original") is True and fixtures.get("remote_resources") is False
        and fixtures.get("calibre", {}).get("path") == calibre and fixtures["calibre"].get("version") == "9.15.0"
        and fixtures.get("azw3_generation", {}).get("exit_code") == 0, "Actual offline EPUB/genuine AZW3 provenance absent")
    refs = report.get("real_conversion_references")
    need(isinstance(refs, dict) and set(refs) == {"epub", "azw3"}, "Independent real conversion references absent")
    for fmt, reference in refs.items():
        need(reference.get("process", {}).get("exit_code") == 0 and reference["process"].get("argv", [None])[0] == calibre
            and reference.get("page_count") == (3 if fmt == "epub" else 4) and reference.get("size_bytes", 0) > 0
            and HEX.fullmatch(reference.get("sha256", "")) and len(reference.get("page_content_sha256", [])) == reference["page_count"], "Independent actual conversion did not certify physical PDF content")
