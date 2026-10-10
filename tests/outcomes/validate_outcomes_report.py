"""Strict M3-T01 receipt contract; no Windows passes inferred from IDs alone."""

from pathlib import Path
import json
import re
import importlib.util

spec = importlib.util.spec_from_file_location("wbs_outcomes_interaction_validation", Path(__file__).resolve().parents[2] / "tests/manual/interaction_receipts.py")
interactions = importlib.util.module_from_spec(spec)
spec.loader.exec_module(interactions)
spec = importlib.util.spec_from_file_location("wbs_outcomes_console_validation", Path(__file__).resolve().parents[2] / "tests/manual/current_launchers.py")
console = importlib.util.module_from_spec(spec)
spec.loader.exec_module(console)

ACCEPTANCE = ["AC-050", "AC-051", "AC-052"]
EXPECTED = {"original-failure": 6, "fallback-success": 0, "fallback-failure": 2,
    "fallback-cancel": 130, "startup-dependency": 3, "converter-failure": 4,
    "output-write": 6, "manifest-finalize": 6, "publish-rename": 6, "cleanup-refusal": 6,
    "engine-cancel": 130, "engine-timeout": 130, "conversion-cancel": 130,
    "conversion-timeout": 130, "log-finalize": 6, "engine-ctrlc": 130, "conversion-ctrlc": 130}
ENGINE_FAULTS = {"output-write", "manifest-finalize", "publish-rename", "cleanup-refusal", "engine-cancel", "engine-timeout", "engine-ctrlc"}
INTERRUPTED = {"engine-cancel", "engine-timeout", "conversion-cancel", "conversion-timeout", "engine-ctrlc", "conversion-ctrlc"}
HEX = re.compile(r"[0-9a-f]{64}")
APPLICATION = {"VERSION", "WinBookSplit.ps1", "WinBookSplit.bat", "requirements.txt",
    "engine/WinBookSplit.Paths.ps1", "engine/WinBookSplit.Diagnostics.ps1", "engine/WinBookSplit.Runtime.ps1",
    "engine/WinBookSplit.Process.ps1", "engine/WinBookSplit.Outcomes.json", "engine/winbooksplit_engine.py",
    "engine/winbooksplit_windows.py", "engine/winbooksplit_conversion.py", "engine/winbooksplit_job.py", "engine/WinBookSplit.Logging.ps1",
    "engine/WinBookSplit.Support.ps1", "Export-WinBookSplitDiagnostics.ps1"}
OUTCOMES = {"original-failure": ("read_error", "unreadable_document"),
    "fallback-success": ("success", "split_complete"), "fallback-failure": ("invalid_input", "invalid_start_pages"),
    "fallback-cancel": ("cancelled", "cancelled"), "startup-dependency": ("failed", "runtime_invalid"),
    "converter-failure": ("error", "conversion_failed"), "output-write": ("write_error", "output_write_failed"),
    "manifest-finalize": ("error", "output_manifest_failed"), "publish-rename": ("error", "output_publish_failed"),
    "cleanup-refusal": ("write_error", "output_write_failed"), "engine-cancel": ("cancelled", "processor_cancelled"),
    "engine-timeout": ("timeout", "processor_timeout"), "conversion-cancel": ("cancelled", "processor_cancelled"),
    "conversion-timeout": ("timeout", "conversion_timeout"), "log-finalize": ("incomplete", "console_finalize_failed"),
    "engine-ctrlc": ("cancelled", "processor_cancelled"), "conversion-ctrlc": ("cancelled", "processor_cancelled")}


def need(value, message):
    if not value:
        raise ValueError(message)


def hashes(value):
    return isinstance(value, dict) and value and all(isinstance(key, str) and isinstance(digest, str) and HEX.fullmatch(digest) for key, digest in value.items())


def validate_case(case):
    need(isinstance(case, dict) and case.get("kind") in EXPECTED, "Unknown outcome case")
    kind = case["kind"]
    batch = case.get("batch")
    need(type(batch) is bool, "Actual BAT scope missing")
    prefix = "BAT" if batch else "PS51" if case.get("host_version", "").startswith("5.1.") else "PS7"
    need(case.get("id") == prefix + "-" + kind, "Outcome case ID/actual host mismatch")
    need(all(case.get(key) is True for key in ("passed", "actual_process", "batch_bytes_unchanged",
        "source_neighbor_prior_unchanged", "unrelated_process_alive", "final_completion_truthful",
        "owned_application_outputs_removed", "post_cleanup_source_neighbor_prior_unchanged")), "Outcome omitted actual preservation/cleanup/survival proof")
    need(type(case.get("exit_code")) is int and case["exit_code"] == EXPECTED[kind]
        and case.get("timed_out") is False, "Final operation native exit differs or harness timeout")
    outcome = case.get("outcome")
    need(isinstance(outcome, dict) and outcome.get("protocol") == "winbooksplit.outcome"
        and type(outcome.get("version")) is int and outcome["version"] == 1
        and type(outcome.get("exit_code")) is int and outcome["exit_code"] == EXPECTED[kind]
        and (outcome.get("status"), outcome.get("code")) == OUTCOMES[kind]
        and type(outcome.get("written_count")) is int, "Final structured outcome incomplete")
    stdout = case.get("stdout")
    need(isinstance(stdout, str), "Actual stdout absent")
    final_lines = [line[len("[OUTCOME] "):] for line in stdout.splitlines() if line.startswith("[OUTCOME] ")]
    need(len(final_lines) == 1, "Actual stdout omitted/duplicated final outcome")
    try:
        printed_outcome = json.loads(final_lines[0])
    except (TypeError, ValueError) as error:
        raise ValueError("Actual stdout final outcome malformed") from error
    need(printed_outcome == outcome, "Stored outcome differs from actual stdout final record")
    success = kind == "fallback-success"
    need((outcome["status"] == "success") is success and ("Done." in case.get("stdout", "")) is success
        and ("Output: " in case.get("stdout", "")) is success, "Output completion disagrees with final structured/native result")
    need((case.get("final_publication") is not None) is (kind in {"fallback-success", "log-finalize"})
        and case.get("final_publication_validated") is (kind in {"fallback-success", "log-finalize"}), "Failed operation published or promised publication was not checked")
    if success:
        need(outcome["written_count"] == 3 and [page for item in case["final_publication"] for page in item.get("page_ids", [])] == [1, 2, 3], "Fallback success lacks complete physical pages")
    else:
        need(outcome["written_count"] == 0, "Failure/incomplete reports successful writes")
    if kind in {"fallback-success", "log-finalize"}:
        engine = outcome.get("engine_result")
        execution = engine.get("execution") if isinstance(engine, dict) else None
        publication = case["final_publication"]
        manifest = case.get("publication_manifest")
        need(isinstance(engine, dict) and engine.get("protocol") == "winbooksplit.result" and engine.get("version") == 1
            and engine.get("status") == "success" and engine.get("code") == "split_complete" and engine.get("exit_code") == 0
            and engine.get("written_count") == 3 and isinstance(execution, dict) and execution.get("written_count") == 3
            and outcome.get("final_directory") == execution.get("final_directory")
            and Path(execution.get("final_directory", "")).parent == Path(case.get("output_base", "")), "Retained publication execution/source binding missing")
        need(isinstance(publication, list) and len(publication) == 3
            and [page for item in publication for page in item.get("page_ids", [])] == [1, 2, 3]
            and [item.get("range") for item in publication] == [[0, 1], [1, 2], [2, 3]]
            and all(item.get("page_count") == 1 and isinstance(item.get("sha256"), str) and HEX.fullmatch(item["sha256"])
                and item.get("size_bytes", 0) > 0 for item in publication), "Publication lacks reopened positive exact physical slices")
        need([item.get("filename") for item in publication] == [item.get("filename") for item in execution.get("outputs", [])]
            and all(item.get("sha256") == planned.get("sha256") and item.get("size_bytes") == planned.get("size_bytes")
                and item.get("page_count") == planned.get("page_count") for item, planned in zip(publication, execution.get("outputs", []))),
            "Actual publication filename/size/hash/count differs from execution")
        content = case.get("source_page_content_sha256")
        need(isinstance(content, list) and len(content) == 3 and all(isinstance(item, str) and HEX.fullmatch(item) for item in content)
            and content == case.get("output_page_content_sha256"), "Publication source/slice physical content mismatch")
        need(isinstance(manifest, dict) and manifest == execution.get("manifest") and manifest.get("status") == "complete"
            and manifest.get("written_count") == 3 and manifest.get("outputs") == execution.get("outputs")
            and isinstance(case.get("publication_manifest_sha256"), str) and HEX.fullmatch(case["publication_manifest_sha256"]), "Actual complete manifest/source digest absent")
    else:
        need(outcome.get("final_directory") is None and case.get("publication_manifest") is None
            and case.get("publication_manifest_sha256") is None and case.get("source_page_content_sha256") is None
            and case.get("output_page_content_sha256") is None, "Failed operation leaked hidden successful publication")
    if kind == "log-finalize":
        need(outcome["status"] == "incomplete" and outcome["code"] == "console_finalize_failed"
            and case.get("log_finalizer_failure_visible") is True
            and outcome.get("engine_result", {}).get("status") == "success", "Finalize failure lost incomplete result or retained publication")
    if kind in (INTERRUPTED - {"engine-timeout", "conversion-timeout"}) | {"fallback-cancel"}:
        need(outcome["status"] == "cancelled", "Cancelled/timeout result not explicit")
    elif kind in {"engine-timeout", "conversion-timeout"}:
        need(outcome["status"] == "timeout", "Timeout does not distinguish deadline reason")
    rows = case.get("engine_records")
    transports = case.get("process_summaries")
    need(isinstance(rows, list) and isinstance(transports, list), "Operation attempt evidence missing")
    need(all(isinstance(frame, dict) and frame.get("protocol") == "winbooksplit.result" and type(frame.get("version")) is int
        and frame["version"] == 1 for frame in rows), "Actual terminal engine protocol/version differs")
    abrupt = {"engine-cancel", "engine-timeout", "conversion-cancel", "engine-ctrlc", "conversion-ctrlc"}
    ordinary = set(EXPECTED) - abrupt - {"startup-dependency", "fallback-success", "fallback-failure", "fallback-cancel"}
    if kind in ordinary:
        engine = outcome.get("engine_result")
        need(isinstance(engine, dict) and engine.get("protocol") == "winbooksplit.result" and engine.get("version") == 1
            and rows == [engine] and len(transports) == 1 and transports[0].get("ExitCode") == engine.get("exit_code")
            and transports[0].get("ResultRecordCount") == 1, "Ordinary outcome lacks exact logged engine frame/native binding")
        if kind != "log-finalize":
            need(all(engine.get(field) == outcome.get(field) for field in ("status", "code", "exit_code", "mode")),
                "Final ordinary outcome differs from last validated engine attempt")
    elif kind in abrupt:
        need(rows == [] and outcome.get("engine_result") is None and len(transports) == 1
            and transports[0].get("ResultRecordCount") == 0 and transports[0].get("ExitCode") == 1,
            "Abrupt interruption fabricated a completed engine result")
    if kind == "startup-dependency":
        need(not rows and not transports and case.get("log_sha256") is None and outcome.get("engine_result") is None
            and case.get("output_members_before_cleanup") == ["prior-output.pdf"], "Dependency refusal touched operation output")
    else:
        need(transports and isinstance(case.get("log_sha256"), str) and HEX.fullmatch(case["log_sha256"]), "Actual native/log observations absent")
        console.validate_console_evidence(case.get("console_evidence"), outcome, case.get("log_sha256"), allow_unfinalized=kind == "log-finalize")
        for value in transports:
            need(isinstance(value, dict) and all(value.get(field) is True for field in ("ParentStopped", "DescendantsStopped", "StreamsComplete", "JobAssigned"))
                and type(value.get("Pid")) is int and value["Pid"] > 0, "Native owned-tree/EOF proof missing")
    if kind.startswith("fallback-"):
        events = case.get("interaction", {}).get("requests", [])
        need(events and events[0].get("stage") == "no_plan", "Actual nonterminal fallback request absent")
        initial = events[0].get("result", {})
        need(len(rows) == len(transports) == 1 and initial.get("protocol") == "winbooksplit.result" and type(initial.get("version")) is int
            and initial["version"] == 1 and initial.get("mode") == "1" and initial.get("status") == "no_plan"
            and initial.get("code") == "no_bookmarks" and initial.get("exit_code") == 5
            and initial.get("written_count") == 0 and initial.get("execution") is None
            and initial.get("fallback_modes") == ["manual"], "Actual initial fallback no-plan missing")
        if kind != "fallback-cancel":
            need(rows[0] == outcome.get("engine_result") and rows[0].get("mode") == "manual"
                and rows[0].get("exit_code") == EXPECTED[kind]
                and all(rows[0].get(field) == outcome.get(field) for field in ("status", "code", "exit_code", "mode")),
                "Outcome selected initial failure rather than final attempted fallback")
        else:
            need(outcome.get("engine_result") == rows[0] == initial, "Fallback cancellation lost actual initial failure")
        need(all(process.get("ExitCode") == frame.get("exit_code") and process.get("ResultRecordCount") == 1
            for process, frame in zip(transports, rows)), "Fallback native/result frame binding differs")
    if kind != "startup-dependency":
        if not kind.startswith("fallback-") and kind != "log-finalize":
            interactions.require_noninteractive(case.get("interaction"), transports)
            need(case.get("stdin_utf8") == "", "Noninteractive fault control supplied fabricated consent")
        else:
            stages = ["no_plan"] if kind == "fallback-cancel" else ["no_plan", "input_ready"] if kind == "fallback-failure" else \
                ["no_plan", "input_ready", "plan_ready"] if kind == "fallback-success" else ["input_ready", "plan_ready"]
            actions = ["cancel"] if kind == "fallback-cancel" else ["retry", "starts"] if kind == "fallback-failure" else \
                ["retry", "starts", "execute"] if kind == "fallback-success" else ["starts", "execute"]
            interactions.validate(case.get("interaction"), transports, stages, actions,
                starts="2,no" if kind == "fallback-failure" else "2,3", execution=outcome.get("engine_result", {}).get("execution"))
    if kind in INTERRUPTED:
        observations = case.get("pid_observations")
        need(isinstance(observations, list) and len(observations) == 2 and len({row.get("pid") for row in observations}) == 2
            and all(type(row.get("pid")) is int and row["pid"] > 0 and row.get("running") is False for row in observations), "Actual authored parent/descendant stop missing")
        need(isinstance(case.get("fault_receipt"), dict) and case["fault_receipt"].get("pid") == observations[0]["pid"],
            "Stopped authored parent differs from readiness receipt PID")
        if kind in {"engine-cancel", "conversion-cancel", "engine-ctrlc", "conversion-ctrlc"}:
            need(transports[-1].get("Cancelled") is True and transports[-1].get("TimedOut") is False, "Token cancellation confused with timeout")
        elif kind == "engine-timeout":
            need(transports[-1].get("TimedOut") is True and transports[-1].get("Cancelled") is False, "Engine deadline not exercised")
        else:
            need(outcome.get("engine_result", {}).get("code") == "conversion_timeout", "Inner owned converter deadline not exercised")
        if kind.endswith("ctrlc"):
            signal = case.get("console_signal")
            need(isinstance(signal, dict) and signal.get("actual_native_console_signal") is True
                and signal.get("hidden_private_console") is True and signal.get("signal_sent") is True
                and signal.get("forced_target_stop") is False and signal.get("controller_ignore_restored") is True,
                "OS Ctrl+C control lacked actual private-console signal/controller restore")
            if signal.get("cmd_prompt_decision") is not None:
                need(case.get("batch") is True and signal["cmd_prompt_decision"] == "N then Enter"
                    and signal.get("cmd_prompt_observed") is True and signal.get("cmd_prompt_after_cancelled_outcome") is True
                    and signal.get("cmd_prompt_owned_pids_stopped") is True, "CMD prompt response preceded actual owned stop/outcome")
    need(hashes(case.get("application_source_sha256")) and hashes(case.get("controlled_application_sha256")), "Shipped/copied byte provenance absent")
    original, controlled = case["application_source_sha256"], case["controlled_application_sha256"]
    allowed = {"WinBookSplit.ps1", "engine/winbooksplit_engine.py"}
    need(set(original) <= set(controlled) and all(controlled[name] == value for name, value in original.items() if name not in allowed), "Acceptance changed an unrelated shipped dependency")
    if kind in ENGINE_FAULTS:
        need(controlled.get("engine/saved_engine.py") == original.get("engine/winbooksplit_engine.py") and isinstance(case.get("fault_receipt"), dict), "Fault adapter did not preserve actual engine bytes/control")
    else:
        need(controlled.get("engine/winbooksplit_engine.py") == original.get("engine/winbooksplit_engine.py"), "Ordinary case substituted engine")
    if kind == "cleanup-refusal":
        diagnostic = outcome.get("engine_result", {}).get("diagnostic")
        need(isinstance(diagnostic, dict) and diagnostic.get("cleanup_complete") is False and diagnostic.get("cleanup_error")
            and case.get("retained_staging") == diagnostic.get("retained_staging")
            and "authored-unrelated.txt" in case["fault_receipt"].get("members", {}), "Unsafe cleanup silently deleted unrelated stage member")
    if kind in {"output-write", "manifest-finalize", "publish-rename", "cleanup-refusal", "converter-failure", "conversion-timeout"}:
        engine = outcome.get("engine_result")
        diagnostic = engine.get("diagnostic") if isinstance(engine, dict) else None
        record = case.get("failure_diagnostic_record")
        need(isinstance(diagnostic, dict) and isinstance(record, dict) and record.get("code") == outcome["code"]
            and record.get("status") in {"failed", "timeout"} and record.get("cleanup_complete") is (kind != "cleanup-refusal")
            and isinstance(case.get("failure_diagnostic_sha256"), str) and HEX.fullmatch(case["failure_diagnostic_sha256"]),
            "Ordinary failure omitted actual recoverable diagnostic or primary result")
    if kind in {"converter-failure", "conversion-timeout"}:
        converted = outcome.get("engine_result", {}).get("diagnostic", {}).get("conversion")
        need(isinstance(converted, dict) and "AUTHORED_CONVERTER_FINAL_STDERR" in converted.get("stderr_tail", "")
            and isinstance(converted.get("argv"), list) and len(converted["argv"]) == 5
            and converted["argv"][0] == case.get("converter_path") and converted["argv"][2] == case.get("fault_receipt", {}).get("output")
            and converted["argv"][3:] == ["--output-profile", "tablet"], "Actual native converter argv/final stderr missing")
        if kind == "converter-failure":
            need(converted.get("exit_code") == 17, "Native converter failure exit overwritten")
    retained_expected = kind in {"cleanup-refusal", "engine-cancel", "engine-timeout", "conversion-cancel", "engine-ctrlc", "conversion-ctrlc"}
    need((case.get("retained_staging") is not None) is retained_expected
        and case.get("known_stage_cleanup_complete") is (not retained_expected), "Stage retention contradicts known ownership/interrupt case")
    if retained_expected:
        retained = Path(case["retained_staging"])
        need(retained.is_absolute() and retained.parent == Path(case.get("output_base", ""))
            and re.fullmatch(r"\.WinBookSplit-stage-[0-9a-f]{32}", retained.name),
            "Retained stage escaped this output base or lacks an exact owned name")
    originals = case.get("source_observations")
    need(isinstance(originals, dict) and len(originals) == 3
        and originals == case.get("after_operation_source_observations") == case.get("post_cleanup_source_observations")
        and all(Path(path).is_absolute() and isinstance(value, dict) and isinstance(value.get("sha256"), str)
            and HEX.fullmatch(value["sha256"]) and type(value.get("size_bytes")) is int and value["size_bytes"] > 0
            and all(type(value.get(field)) is int for field in ("device", "inode", "attributes")) for path, value in originals.items()),
        "Actual source/neighbor/prior identity and hash comparison missing")
    need(case.get("output_members_after_cleanup") == ["prior-output.pdf"]
        and isinstance(case.get("output_members_before_cleanup"), list)
        and "prior-output.pdf" in case["output_members_before_cleanup"]
        and len(set(case["output_members_before_cleanup"])) == len(case["output_members_before_cleanup"]), "Known flat cleanup membership evidence missing")
    need(type(case.get("unrelated_pid")) is int and case["unrelated_pid"] > 0
        and isinstance(case.get("command"), list) and case["command"] and isinstance(case.get("cwd"), str)
        and Path(case["cwd"]).is_absolute() and Path(case.get("output_base", "")).is_absolute(), "Actual native command/path/survival receipt incomplete")


def validate_outcomes_report(report, shells):
    need(isinstance(report, dict) and report.get("schema_version") == 1 and report.get("task_id") == "M3-T01"
        and report.get("result") == "OUTCOME_REGRESSION_PASSED" and report.get("success") is True
        and type(report.get("exit_code")) is int and report["exit_code"] == 0 and report.get("acceptance_ids") == ACCEPTANCE,
        "Outcome acceptance report/header incomplete")
    need(all(report.get(field) is True for field in ("source_unchanged", "machine_settings_unchanged", "baseline_guards_preserved", "owned_temp_removed")), "Outcome report lacks source/settings/cleanup proofs")
    hosts = report.get("host_cases")
    need(isinstance(hosts, list) and len(hosts) == 2 and {row.get("id") for row in hosts} == {"PS51", "PS7"}
        and {Path(row.get("shell_executable", "")) for row in hosts} == {Path(shell) for shell in shells}, "Both requested actual supported hosts required")
    for host in hosts:
        need(host.get("stored_policies") == host.get("policies_after") and host.get("syntax_error_count") == 0, "Actual host stored policies/syntax differ")
    cases = report.get("cases")
    needed = {host + "-" + kind for host in ("PS51", "PS7", "BAT") for kind in EXPECTED}
    need(isinstance(cases, list) and len(cases) == len(needed) and report.get("case_count") == len(needed)
        and {row.get("id") for row in cases} == needed, "Missing/duplicate actual PS51/PS7/BAT outcome case")
    by_host = {host["id"]: host for host in hosts}
    for case in cases:
        validate_case(case)
        host = by_host["PS51" if case["batch"] else case["id"].split("-", 1)[0]]
        need(Path(case["shell_executable"]) == Path(host["shell_executable"]) and case["host_version"] == host["host_version"], "Case ran a different host")
        need(set(case["application_source_sha256"]) == APPLICATION
            and case["application_source_sha256"] == {name: report.get("tested_path_sha256", {}).get(name) for name in APPLICATION},
            "Actual copied application differs from report tested source bytes")
        if case["kind"] in ENGINE_FAULTS:
            need(case["controlled_application_sha256"]["engine/winbooksplit_engine.py"]
                == report.get("tested_path_sha256", {}).get("tests/outcomes/fault_engine.py"),
                "Controlled engine differs from exact tested fault adapter bytes")
        if kind := case.get("kind"):
            if kind in {"converter-failure", "conversion-cancel", "conversion-timeout", "conversion-ctrlc"}:
                need(case.get("converter_sha256") == report.get("converter_compilation", {}).get("executable_sha256"),
                    "Actual converter executable differs from recorded compiled native fixture")
    unrelated = report.get("unrelated_process_observation")
    need(isinstance(unrelated, dict) and unrelated.get("alive_after_all_operations") is True
        and unrelated.get("running_after_retained_handle_stop") is False and type(unrelated.get("retained_pid")) is int
        and unrelated.get("actual_identity", {}).get("pid") == unrelated["retained_pid"]
        and all(case["unrelated_pid"] == unrelated["retained_pid"] for case in cases), "Unrelated actual self-PID/readiness/lifetime missing")
    need(hashes(report.get("tested_path_sha256")), "Actual tested source-byte digest absent")
    need({"tests/outcomes/characterize_outcomes.py", "tests/outcomes/validate_outcomes_report.py",
        "tests/outcomes/fault_engine.py", "tests/outcomes/OwnedConverter.cs"} <= set(report["tested_path_sha256"]),
        "Actual harness/validator/fault/native source provenance absent")
    compiled = report.get("converter_compilation")
    need(isinstance(compiled, dict) and compiled.get("exit_code") == 0
        and compiled.get("source_sha256") == report["tested_path_sha256"].get("tests/outcomes/OwnedConverter.cs")
        and isinstance(compiled.get("executable_sha256"), str) and HEX.fullmatch(compiled["executable_sha256"]),
        "Native converter compilation/source provenance differs")
