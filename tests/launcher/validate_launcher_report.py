"""Strict native launcher/menu evidence; this does not certify Explorer use."""

from hashlib import sha256
import importlib.util
import json
import math
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("wbs_launcher_cli_validation", ROOT / "tests/cli/validate_cli_report.py")
cli = importlib.util.module_from_spec(spec)
spec.loader.exec_module(cli)
need, hashes, identities, exact_cases, HEX = cli.need, cli.hashes, cli.identities, cli.exact_cases, cli.HEX
APPLICATION = cli.APPLICATION
ACCEPTANCE = ["AC-058", "AC-059"]
COMMON = ["source-pdf", "source-quoted-pdf", "source-epub", "source-azw3", "source-blank", "source-c", "source-trimmed-c",
          "menu-level1", "menu-level2", "menu-manual", "menu-cancel", "menu-eof", "menu-invalid-level1", "menu-invalid-manual",
          "fallback-invalid-manual", "fallback-y", "fallback-level1", "fallback-n", "fallback-c", "fallback-eof", "fallback-invalid-cancel"]
BAT_ONLY = ["drop-pdf", "drop-epub", "drop-azw3", "source-delayed-expansion", "two-inputs", "three-inputs", "empty-second-third"]
EARLY = {"source-blank", "source-c", "source-trimmed-c", "two-inputs", "three-inputs"}
CANCELLED = {"source-blank", "source-c", "source-trimmed-c", "menu-cancel", "menu-eof", "fallback-n", "fallback-c", "fallback-eof", "fallback-invalid-cancel"}
EXTRA = {"two-inputs", "three-inputs", "empty-second-third"}
FALLBACK = {name for name in COMMON if name.startswith("fallback-")}
INVALID_MENU = ["", "maybe", "yes", "zebra", "Mno", "1oops"]
INVALID_FALLBACK = ["maybe", "yes", "", "2", "1oops"]
SENTINEL_SOURCE = "param([string]$InputFile)\n[IO.File]::WriteAllText($env:WBS_LAUNCHER_STARTED, 'unexpected application launch')\nexit 73\n"


def recipe(kind, source):
    if kind in {"source-pdf", "source-epub", "source-azw3", "source-delayed-expansion"}:
        return str(source) + "\n1\n\n", "1"
    if kind == "source-quoted-pdf":
        return '"' + str(source) + '"\n1\n\n', "1"
    if kind == "source-blank":
        return "\n", None
    if kind == "source-c":
        return "C\n", None
    if kind == "source-trimmed-c":
        return "  c  \n", None
    if kind in EXTRA:
        return "", None
    if kind in {"menu-level1", "drop-pdf", "drop-epub", "drop-azw3"}:
        return " 1 \n\n", "1"
    if kind == "menu-level2":
        return " 2 \n\n", "2"
    if kind == "menu-manual":
        return " m \n1,3,5\n\n", "manual"
    if kind == "menu-cancel":
        return " c \n", None
    if kind == "menu-eof":
        return "", None
    if kind == "menu-invalid-level1":
        return "\n".join(INVALID_MENU + [" 1 ", "", ""]), "1"
    if kind == "menu-invalid-manual":
        return "manual\nmango\n m \n1,3,5\n\n", "manual"
    if kind == "fallback-invalid-manual":
        return "2\n" + "\n".join(INVALID_FALLBACK + [" m ", "1,3,5", "", ""]), "manual"
    if kind == "fallback-y":
        return "1\n y \n1,3,5\n\n", "manual"
    if kind == "fallback-level1":
        return "2\n 1 \n\n", "1"
    if kind == "fallback-n":
        return "1\n n \n", "1"
    if kind == "fallback-c":
        return "1\n c \n", "1"
    if kind == "fallback-eof":
        return "1\n", "1"
    if kind == "fallback-invalid-cancel":
        return "2\n" + "\n".join(INVALID_FALLBACK + [" c ", ""]), "2"
    raise ValueError("Unknown native launcher case")


def validate_case(row):
    need(isinstance(row, dict) and row.get("kind") in set(COMMON + BAT_ONLY), "Unknown launcher case")
    kind, host = row["kind"], row.get("host_id")
    need(host in {"PS51", "PS7", "BAT"} and row.get("id") == host + "-" + kind
         and (host == "BAT" or kind in COMMON), "Launcher case/host ID differs")
    need(isinstance(row.get("host_version"), str) and row["host_version"].startswith("7." if host == "PS7" else "5.1."), "Actual launcher host version absent")
    expected_exit = 2 if kind in EXTRA else 130 if kind in CANCELLED else 0
    need(type(row.get("exit_code")) is int and row["exit_code"] == expected_exit and row.get("timed_out") is False
         and type(row.get("pid")) is int and row["pid"] > 0 and row.get("stdin") == "PIPE_UTF8"
         and type(row.get("elapsed_seconds")) in (int, float) and math.isfinite(row["elapsed_seconds"])
         and 0 <= row["elapsed_seconds"] < row.get("timeout_seconds", 0), "Actual native launcher status/input/deadline absent")
    need(all(row.get(name) is True for name in ("passed", "actual_process", "source_read_only_observed", "source_neighbor_prior_unchanged",
         "application_unchanged", "known_output_cleanup_complete")), "Launcher preservation/cleanup proof absent")
    for name in ("stdout", "stderr"):
        need(isinstance(row.get(name), str) and row.get(name + "_sha256") == sha256(row[name].encode("utf-8")).hexdigest(), "Actual launcher stream/hash differs")
    if kind in CANCELLED:
        need(not any(label in row["stdout"] + row["stderr"] for label in ("[!] Error:", "[!] Processor failure:", "Done.")), "Cancellation was mislabeled as failure or success")
    answers, expected_mode = recipe(kind, row.get("input_path"))
    need(row.get("stdin_utf8") == answers and row.get("stdin_sha256") == sha256(answers.encode("utf-8")).hexdigest(), "Actual exact menu/source input differs")
    for field in ("source_observations_before", "source_observations_after", "source_observations_after_cleanup"):
        identities(row.get(field))
    need(row["source_observations_before"] == row["source_observations_after"] == row["source_observations_after_cleanup"]
         and row["source_observations_before"]["source"]["attributes"] & 1, "Source/neighbor/prior changed or read-only guard absent")
    extra = row.get("additional_inputs_before")
    need(isinstance(extra, dict) and len(extra) == 2 and extra == row.get("additional_inputs_after")
         and all(isinstance(item, dict) and HEX.fullmatch(item.get("sha256", "")) and item.get("size_bytes", 0) > 0
                 and type(item.get("inode")) is int and item["inode"] > 0 for item in extra.values()), "Additional dropped source identities changed or absent")
    need(row.get("output_members_before") == row.get("output_members_after_cleanup") == ["prior-output.pdf"], "Launcher cleanup touched prior output")
    need(hashes(row.get("application_sha256")) and set(row["application_sha256"]) == APPLICATION
         and hashes(row.get("actual_application_sha256")) and set(row["actual_application_sha256"]) == APPLICATION,
         "Exact shipped/copied launcher application hashes absent")
    original, actual = row["application_sha256"], row["actual_application_sha256"]
    need(all(actual[name] == value for name, value in original.items() if name != "WinBookSplit.ps1"), "BAT or shipped helpers were modified")
    command, parameters = row.get("command"), row.get("parameters")
    need(isinstance(command, list) and command and isinstance(parameters, list) and all(isinstance(arg, str) for arg in command + parameters)
         and isinstance(row.get("cwd"), str) and row["cwd"] != str(Path(row["input_path"]).parent), "Literal command/unrelated CWD absent")
    if host == "BAT":
        need(Path(command[0]).name.lower() == "cmd.exe" and command[1:4] == ["/d", "/v:on" if kind == "source-delayed-expansion" else "/v:off", "/c"]
             and len(command) == 5 and row.get("wrapper_sha256") == sha256(row.get("wrapper_text", "").encode("ascii")).hexdigest(), "Native BAT wrapper bytes/argv absent")
        count = 2 if kind == "two-inputs" else 3 if kind in {"three-inputs", "empty-second-third"} else 0 if kind.startswith("source-") else 1
        expected_wrapper = '@echo off\r\nsetlocal ' + ('EnableDelayedExpansion' if kind == "source-delayed-expansion" else 'DisableDelayedExpansion') + '\r\n"%WBS_LAUNCHER_BAT%"' + ''.join(' "%WBS_LAUNCHER_INPUT' + str(index) + '%"' for index in range(1, count + 1)) + '\r\n'
        need(row["wrapper_text"] == expected_wrapper and len(parameters) == count and (count == 0 or parameters[0] == row["input_path"]), "BAT argument count or literal wrapper differs")
        expected_mods = ["copied PS launch sentinel only"] if kind in EXTRA else ["copied parameter default OutputDirectory", "copied parameter default PythonPath", "copied parameter default CalibrePath", "copied NoPause default true"]
        need(row.get("modifications") == expected_mods, "Undeclared copied BAT dependency control")
        if kind in EXTRA:
            need(actual["WinBookSplit.ps1"] == sha256(SENTINEL_SOURCE.encode("utf-8-sig")).hexdigest()
                 and row.get("powershell_launch_sentinel_absent") is True, "Multi-input rejection launched PowerShell or lacks sentinel control")
    else:
        need(command[0] == row.get("shell_executable") and "-NoProfile" in command and "-File" in command
             and "-NonInteractive" not in command and command[command.index("-File") + 2:] == parameters
             and Path(command[command.index("-File") + 1]).name == "WinBookSplit.ps1"
             and original == actual and row.get("modifications") == [], "Direct launcher source/arguments changed")
        need("-NoPause" in parameters and "-PythonPath" in parameters and "-OutputDirectory" in parameters,
             "Native launcher isolated dependencies/output absent")
        need(("-InputFile" not in parameters) == kind.startswith("source-"), "Source selection did not use actual absent input")
    final, rows, transports = row.get("outcome"), row.get("engine_records"), row.get("process_summaries")
    if kind in EXTRA:
        need(final is None and rows == transports == [] and row.get("log_sha256") is None
             and row.get("console_operation_outcomes") == []
             and row.get("output_members_after") == ["prior-output.pdf"]
             and all(text in row["stdout"] for text in ("one", "PDF", "EPUB", "AZW3")), "BAT extra inputs were not refused before processing")
        return
    need(isinstance(final, dict) and final.get("protocol") == "winbooksplit.outcome" and type(final.get("version")) is int
         and final["version"] == 1 and type(final.get("exit_code")) is int and final["exit_code"] == expected_exit
         and type(final.get("written_count")) is int, "Final native launcher outcome absent")
    printed = [json.loads(line[len("[OUTCOME] "):]) for line in row["stdout"].splitlines() if line.startswith("[OUTCOME] ")]
    need(printed == [final], "Native final outcome is missing/duplicate or differs")
    # Read-Host prompt text is a console host UI message, omitted by actual
    # redirected PS5.1 streams. Absent InputFile + exact authored stdin +
    # resulting immutable source identity prove this native path. Visible
    # wording is separately checked by helper tests and human console evidence.
    if kind in EARLY:
        need(final.get("status") == "cancelled" and final.get("code") == "cancelled" and final["written_count"] == 0
             and final.get("engine_result") is None and final.get("final_directory") is None and rows == transports == []
             and row.get("log_sha256") is None and row.get("output_members_after") == ["prior-output.pdf"]
             and row.get("console_operation_outcomes") == []
             and "[DEPENDENCY]" not in row["stdout"] and "Done." not in row["stdout"], "Source cancellation wrote/probed/processed or falsely succeeded")
        return
    scope = row.get("console_log_scope")
    need(isinstance(scope, dict) and row.get("console_operation_outcomes") == [final]
         and scope.get("sha256") == row.get("log_sha256") and HEX.fullmatch(scope.get("sha256", ""))
         and scope.get("members") == [".WinBookSplit-console-owner.json", "console.log"]
         and Path(scope.get("path", "")).parent.parent == Path(row["output_base"])
         and scope.get("owner") == {"run_id": Path(scope["path"]).parent.name.removeprefix(".WinBookSplit-console-"), "kind": "console"}
         and re.fullmatch(r"\.WinBookSplit-console-[0-9a-f]{32}", Path(scope["path"]).parent.name), "Exact owned log/final operation footer proof absent")
    if kind in {"menu-cancel", "menu-eof"}:
        need(final.get("status") == "cancelled" and final.get("code") == "cancelled" and final["written_count"] == 0
             and final.get("engine_result") is None and final.get("final_directory") is None and rows == transports == []
             and row.get("final_publication") is None and "Done." not in row["stdout"], "Menu cancel performed a split or falsely succeeded")
        return
    expected_attempts = 1 if kind in CANCELLED or kind not in FALLBACK else 2
    need(isinstance(rows, list) and len(rows) == expected_attempts and isinstance(transports, list) and len(transports) == expected_attempts
         and final.get("engine_result") == rows[-1] and isinstance(row.get("log_sha256"), str) and HEX.fullmatch(row["log_sha256"]), "Exact actual engine/native/log attempts absent")
    for frame, transport in zip(rows, transports):
        need(frame.get("protocol") == "winbooksplit.result" and type(frame.get("version")) is int and frame["version"] == 1
             and transport.get("ExitCode") == frame.get("exit_code") and transport.get("ResultRecordCount") == 1
             and all(transport.get(field) is True for field in ("JobAssigned", "ParentStopped", "DescendantsStopped", "StreamsComplete")), "Owned engine stop/stream/result proof absent")
    if kind in FALLBACK:
        first_mode = "2" if kind in {"fallback-invalid-manual", "fallback-level1", "fallback-invalid-cancel"} else "1"
        need(rows[0].get("status") == "no_plan" and rows[0].get("exit_code") == 5 and rows[0].get("written_count") == 0
             and rows[0].get("execution") is None and rows[0].get("mode") == first_mode
             and rows[0].get("code") == ("no_bookmarks_at_level" if first_mode == "2" else "no_bookmarks")
             and rows[0].get("fallback_modes") == (["1", "manual"] if first_mode == "2" else ["manual"])
             and "[FALLBACK]" in row["stdout"], "Fallback did not follow actual non-writing requested no-plan")
    if kind in {"menu-invalid-level1", "menu-invalid-manual"}:
        minimum = len(INVALID_MENU) if kind == "menu-invalid-level1" else 2
        need(row.get("menu_invalid_message_count", 0) >= minimum, "Arbitrary initial choices did not all reprompt")
    if kind in {"fallback-invalid-manual", "fallback-invalid-cancel"}:
        need(row.get("fallback_invalid_message_count", 0) >= len(INVALID_FALLBACK), "Arbitrary fallback choices did not all reprompt")
    if kind in CANCELLED:
        need(final.get("status") == "cancelled" and final.get("code") == "cancelled" and final["written_count"] == 0
             and final.get("mode") == expected_mode and final.get("final_directory") is None
             and row.get("final_publication") is None and "Done." not in row["stdout"], "Fallback cancellation falsely succeeded")
        return
    frame, outputs = rows[-1], row.get("final_publication")
    execution = frame.get("execution")
    need(frame.get("mode") == final.get("mode") == expected_mode and frame.get("status") == final.get("status") == "success"
         and frame.get("code") == final.get("code") == "split_complete" and isinstance(execution, dict), "Selected exact mode did not produce final success")
    bound = execution.get("source_identity") if Path(row["input_path"]).suffix.lower() == ".pdf" else execution.get("original_ebook_identity")
    need(isinstance(bound, dict) and Path(bound.get("path", "")) == Path(row["input_path"])
         and bound.get("sha256") == row["source_observations_before"]["source"]["sha256"]
         and bound.get("size_bytes") == row["source_observations_before"]["source"]["size_bytes"], "Native split did not bind the exact literal original source")
    ranges, content = row.get("expected_ranges"), row.get("expected_content_sha256")
    need(isinstance(ranges, list) and bool(ranges) and isinstance(content, list) and content and all(isinstance(item, str) and HEX.fullmatch(item) for item in content)
         and isinstance(outputs, list) and len(outputs) == len(ranges) and row.get("output_content_sha256") == content
         and final["written_count"] == frame.get("written_count") == execution.get("written_count") == len(outputs), "Exact positive publication/content absent")
    previous = 0
    for pair, item, entry in zip(ranges, outputs, row.get("expected_entries", [])):
        need(isinstance(pair, list) and len(pair) == 2 and all(type(page) is int for page in pair)
             and pair[0] == previous < pair[1] <= len(content) and item.get("range") == pair
             and item.get("filename") == entry.get("filename") and item.get("page_count") == pair[1] - pair[0]
             and item.get("size_bytes", 0) > 0 and isinstance(item.get("sha256"), str) and HEX.fullmatch(item["sha256"]), "Native outputs differ from exact full physical ranges")
        previous = pair[1]
    need(previous == len(content) and len(row.get("expected_entries", [])) == len(outputs)
         and row.get("publication_manifest") == execution.get("manifest") and execution["manifest"].get("outputs") == execution.get("outputs")
         and HEX.fullmatch(row.get("publication_manifest_sha256", "")) and final.get("final_directory") == execution.get("final_directory")
         and Path(final["final_directory"]).parent == Path(row["output_base"]) and "Done." in row["stdout"], "Publication coverage/manifest/final path proof absent")
    if Path(row["input_path"]).suffix.lower() == ".pdf":
        need(ranges == cli.RANGES[expected_mode] and [page for item in outputs for page in item.get("page_ids", [])] == list(range(1, 7)), "Authored PDF physical pages omitted/reordered")
    else:
        conversion = execution.get("conversion")
        need(isinstance(conversion, dict) and conversion.get("generated_pdf_identity", {}).get("page_content_sha256") == content
             and execution.get("original_ebook_identity", {}).get("sha256") == row["source_observations_before"]["source"]["sha256"]
             and row.get("temporary_conversion_absent") is True and execution.get("retained_intermediate") is None,
             "Actual ebook temporary conversion/source/content/cleanup proof absent")


def validate_launcher_report(report, requested_shells, calibre):
    need(isinstance(report, dict) and type(report.get("schema_version")) is int and report["schema_version"] == 1
         and report.get("task_id") == "M3-T03" and report.get("result") == "LAUNCHER_REGRESSION_PASSED"
         and report.get("success") is True and type(report.get("exit_code")) is int and report["exit_code"] == 0, "Launcher report did not pass")
    need(report.get("acceptance_ids") == ACCEPTANCE and report.get("evidence_kind") == "AUTOMATED_NATIVE_SUPPLEMENT"
         and report.get("human_explorer_tested") is False and report.get("not_run") == ["human Explorer/key presses", "release package", "clean OS"], "Automated launcher receipt mislabels human/Explorer evidence")
    need(all(report.get(key) is True for key in ("source_unchanged", "baseline_guards_preserved", "machine_settings_unchanged", "owned_temp_removed")), "Launcher source/settings/cleanup proof absent")
    manifest = report.get("tested_path_sha256")
    need(hashes(manifest) and APPLICATION | {"tests/launcher/characterize_launcher.py", "tests/launcher/validate_launcher_report.py", "tests/run_tests.py"} <= set(manifest), "Launcher tested raw source scope absent")
    hosts = exact_cases(report.get("host_cases"), {"PS51", "PS7"})
    need(len(requested_shells) == 2 and len(set(requested_shells)) == 2 and {h.get("shell_executable") for h in hosts.values()} == set(requested_shells), "Requested native hosts differ")
    for name, host in hosts.items():
        need(host.get("passed") is True and host.get("host_major") == (5 if name == "PS51" else 7) and host.get("syntax_error_count") == 0
             and host.get("stored_policies") == host.get("policies_after") and isinstance(host.get("stored_policies"), list) and host["stored_policies"], "Actual hosts or unchanged policies absent")
    rows = exact_cases(report.get("cases"), {host + "-" + kind for host in hosts for kind in COMMON} | {"BAT-" + kind for kind in COMMON + BAT_ONLY})
    need(type(report.get("case_count")) is int and report["case_count"] == len(rows), "Native launcher count differs")
    for row in rows.values():
        validate_case(row)
        host = hosts["PS51" if row["host_id"] == "BAT" else row["host_id"]]
        need(row["host_version"] == host["host_version"] and row["shell_executable"] == host["shell_executable"], "Case used an unrequested host")
        need(all(manifest[name] == value for name, value in row["application_sha256"].items()), "Copied application differs from tested source bytes")
    fixtures, refs = report.get("fixture_provenance"), report.get("real_conversion_references")
    need(isinstance(fixtures, dict) and fixtures.get("authored_original") is True and fixtures.get("remote_resources") is False
         and fixtures.get("calibre", {}).get("path") == calibre and fixtures["calibre"].get("version") == "9.15.0"
         and fixtures.get("azw3_generation", {}).get("exit_code") == 0, "Offline genuine EPUB/AZW3 provenance absent")
    need(isinstance(refs, dict) and set(refs) == {"epub", "azw3"}, "Independent converter controls absent")
    for fmt, reference in refs.items():
        need(reference.get("process", {}).get("exit_code") == 0 and reference["process"].get("argv", [None])[0] == calibre
             and reference.get("page_count") == (3 if fmt == "epub" else 4) and reference.get("size_bytes", 0) > 0
             and HEX.fullmatch(reference.get("sha256", "")) and len(reference.get("page_content_sha256", [])) == reference["page_count"], "Actual independent ebook physical conversion absent")
