"""Independent strict data validator for native M3-T04 UX receipts."""

from hashlib import sha256
import importlib.util
import json
import math
from pathlib import Path
import re
import unicodedata


ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("wbs_ux_cli_validation", ROOT / "tests/cli/validate_cli_report.py")
cli = importlib.util.module_from_spec(spec)
spec.loader.exec_module(cli)
need, HEX, APPLICATION = cli.need, cli.HEX, cli.APPLICATION
APPLICATION_KINDS = ["nested-confirm", "confirm-cancel", "confirm-blank", "confirm-eof", "confirm-invalid",
                     "manual-physical", "ebook-manual-epub", "ebook-manual-azw3", "fallback-cancel",
                     "fallback-manual", "noninteractive", "preview", "pending-timeout"]
CANCELLED = {"confirm-cancel", "confirm-blank", "confirm-eof", "fallback-cancel", "pending-timeout"}
NESTED_RANGES = [[0, 2], [2, 3], [3, 6], [6, 8], [8, 12]]
PROMPT = "Create these chapter PDFs? [Y] yes, [C] cancel"


def raw_streams(row):
    need(type(row.get("pid")) is int and row["pid"] > 0 and type(row.get("exit_code")) is int
         and row.get("actual_process") is True and row.get("timed_out") is False
         and row.get("streams_complete") is True and isinstance(row.get("command"), list)
         and bool(row["command"]) and all(isinstance(arg, str) for arg in row["command"])
         and isinstance(row.get("cwd"), str), "Concrete actual native process/streams missing")
    elapsed, timeout = row.get("elapsed_seconds"), row.get("timeout_seconds")
    need(type(elapsed) in (int, float) and math.isfinite(elapsed) and type(timeout) in (int, float)
         and 0 <= elapsed < timeout and timeout > 0, "Native bounded elapsed/deadline invalid")
    for name in ("stdout", "stderr"):
        need(isinstance(row.get(name), str) and row.get(name + "_sha256") == sha256(row[name].encode("utf-8")).hexdigest(),
             "Native raw stream/hash differs")


def validate_plan_event(event, mode, base):
    need(isinstance(event, dict) and event.get("protocol") == "winbooksplit.interaction"
         and type(event.get("version")) is int and event["version"] == 1 and event.get("stage") == "plan_ready"
         and event.get("mode") == mode and isinstance(event.get("session"), str)
         and re.fullmatch(r"[0-9a-f]{32}", event["session"]) and type(event.get("sequence")) is int and event["sequence"] > 0,
         "Held plan interaction protocol/session/sequence invalid")
    canonical, plan = event.get("plan_json"), event.get("plan")
    need(isinstance(canonical, str) and isinstance(plan, dict) and json.loads(canonical) == plan
         and canonical == json.dumps(plan, sort_keys=True, ensure_ascii=True, separators=(",", ":"))
         and event.get("plan_sha256") == sha256(canonical.encode("utf-8")).hexdigest(), "Held canonical plan/digest differs")
    need(plan.get("mode") == mode and plan.get("output_naming", {}).get("resolved_base") == base,
         "Held plan mode/destination differs")
    identity = plan.get("source_identity")
    need(isinstance(identity, dict) and identity.get("binding") == "reader_snapshot"
         and isinstance(identity.get("path"), str) and isinstance(identity.get("resolved_path"), str)
         and isinstance(identity.get("sha256"), str) and HEX.fullmatch(identity["sha256"])
         and type(identity.get("size_bytes")) is int and identity["size_bytes"] > 0, "Captured reader identity is invalid")
    for entry in plan.get("entries", []):
        name = entry.get("filename")
        need(isinstance(name, str) and name == name.strip() and name not in {".pdf", "..pdf"}
             and not re.search(r'[<>:"/\\|?*\x00-\x1f]', name)
             and not any(unicodedata.category(char) in {"Cc", "Cf", "Cs"} for char in name),
             "Held plan has an unsafe PDF basename")
    return plan


def expected_ranges(kind):
    if kind == "ebook-manual-epub":
        return [[0, 1], [1, 2], [2, 3]]
    if kind == "ebook-manual-azw3":
        return [[0, 1], [1, 2], [2, 4]]
    return [[0, 4], [4, 8], [8, 12]] if kind in {"manual-physical", "fallback-manual"} else NESTED_RANGES


def expected_entries(kind):
    ranges = expected_ranges(kind)
    if kind in {"manual-physical", "fallback-manual", "ebook-manual-epub", "ebook-manual-azw3"}:
        titles = [f"Section (Page {start + 1}-{end})" for start, end in ranges]
        reasons = ["manual"] * 3
    else:
        titles = ["Front matter", "Parent A Å - Opening pages", "Child A1", "Child A2", "Parent B childless"]
        reasons = ["front_matter", "parent_opening", "bookmark", "bookmark", "parent_fallback"]
    return titles, reasons


def validate_application_case(row):
    need(isinstance(row, dict) and row.get("kind") in APPLICATION_KINDS and row.get("host_id") in {"PS51", "PS7"}, "Unknown UX case")
    kind, host = row["kind"], row["host_id"]
    need(row.get("id") == host + "-" + kind and isinstance(row.get("host_version"), str)
         and row["host_version"].startswith("5.1." if host == "PS51" else "7."), "Native UX host/version ID differs")
    raw_streams(row)
    need(row.get("passed") is True and row.get("application_unchanged") is True and row.get("source_read_only_observed") is True
         and cli.hashes(row.get("application_sha256")) and set(row["application_sha256"]) == APPLICATION,
         "Exact copied application/preservation evidence missing")
    for key in ("source_observations_before", "source_observations_after", "source_observations_after_cleanup"):
        cli.identities(row.get(key))
    need(row["source_observations_before"] == row["source_observations_after"] == row["source_observations_after_cleanup"]
         and row["source_observations_before"]["source"]["attributes"] & 1
         and row.get("output_members_after_cleanup") == ["prior-output.pdf"], "Actual UX changed source/neighbor/prior or cleanup")
    parameters, command = row.get("parameters"), row["command"]
    need(isinstance(parameters, list) and command[0] == row.get("shell_executable") and "-NoProfile" in command
         and "-File" in command and Path(command[command.index("-File") + 1]).name == "WinBookSplit.ps1"
         and command[command.index("-File") + 2:] == parameters and "-InputFile" in parameters
         and parameters[parameters.index("-InputFile") + 1] == row.get("input_path")
         and parameters[parameters.index("-OutputDirectory") + 1] == row.get("output_base"), "Literal native UX argv differs")
    unavailable = kind in {"noninteractive", "preview"}
    need(("-NonInteractive" in command[:command.index("-File")]) == unavailable,
         "Native host stdin/prompt control changed")
    need(("-NonInteractive" in parameters) == (kind == "noninteractive") and ("-Preview" in parameters) == (kind == "preview"),
         "Application no-prompt control changed")
    expected_input = {"confirm-cancel": "C\n", "confirm-blank": "\n", "confirm-eof": "",
                      "confirm-invalid": "Q\nmaybe\nY\nN\n", "manual-physical": "M\n1,5,9\nY\nN\n",
                      "ebook-manual-epub": "M\nN\n1,2,3\nY\nN\n", "ebook-manual-azw3": "M\nN\n1,2,3\nY\nN\n",
                      "fallback-cancel": "C\n", "fallback-manual": "M\n1,5,9\nY\nN\n",
                      "noninteractive": "", "preview": "", "pending-timeout": ""}.get(kind, "Y\nN\n")
    need(row.get("stdin_utf8") == expected_input, "Actual authored consent/manual input differs")
    expected_exit = 130 if kind in CANCELLED else 0
    final = row.get("outcome")
    need(isinstance(final, dict) and final.get("protocol") == "winbooksplit.outcome" and type(final.get("version")) is int
         and final["version"] == 1 and final.get("exit_code") == row["exit_code"] == expected_exit,
         "Final/native UX outcome differs")
    need(type(final.get("exit_code")) is int and type(final.get("written_count")) is int, "Typed final outcome counts missing")
    need([json.loads(line[len("[OUTCOME] "):]) for line in row["stdout"].splitlines() if line.startswith("[OUTCOME] ")] == [final],
         "Missing/duplicate/conflicting final UX stdout frame")
    text = row.get("console_log")
    if kind == "preview":
        need(text is None and row.get("console_log_sha256") is None and row.get("terminal_evidence_source") == "stdout_outcome"
             and all(row.get(name) == [] for name in ("engine_records", "process_summaries", "plan_events", "interaction_replies", "interaction_events"))
             and "Log: " not in row["stdout"], "Preview fabricated a persisted console/transport record")
        need(isinstance(final.get("engine_result"), dict), "Preview stdout outcome omitted its strict engine result")
        # Preview's StringWriter is deliberately unpersisted. Bind the only
        # available terminal result to authoritative stdout; claim no child log.
        rows, plans, replies, interactions, process = [final["engine_result"]], [], [], [], None
    else:
        need(isinstance(text, str) and row.get("console_log_sha256") == sha256(text.encode("utf-8")).hexdigest()
             and row.get("terminal_evidence_source", "console_log") == "console_log", "Actual console log/hash differs")
        plans = [json.loads(line[len("[PLAN] "):]) for line in text.splitlines() if line.startswith("[PLAN] ")]
        replies = [json.loads(line[len("[INTERACTION-REPLY] "):]) for line in text.splitlines() if line.startswith("[INTERACTION-REPLY] ")]
        interactions = [json.loads(line[len("[WBS-INTERACTION] "):]) for line in text.splitlines() if line.startswith("[WBS-INTERACTION] ")]
        rows = [json.loads(line) for line in text.splitlines() if line.startswith("{")]
        processes = [json.loads(line[len("[PROCESS] "):]) for line in text.splitlines() if line.startswith("[PROCESS] ")]
        need(plans == row.get("plan_events") and replies == row.get("interaction_replies") and interactions == row.get("interaction_events") and rows == row.get("engine_records")
             and processes == row.get("process_summaries") and len(processes) == 1
             and [json.loads(line[len("[OPERATION-OUTCOME] "):]) for line in text.splitlines() if line.startswith("[OPERATION-OUTCOME] ")] == [final],
             "Actual console events/result/transport/footer differ")
        process = processes[0]
        need(all(process.get(key) is True for key in ("JobAssigned", "ParentStopped", "DescendantsStopped", "StreamsComplete"))
             and not any(process.get(key) for key in ("StartError", "StopError", "StreamError", "ResultError")),
             "Actual engine job/tree/EOF proof incomplete")
    if not unavailable:
        need(process.get("InputWriterStopped") is True and not process.get("InteractionError") and not process.get("InputError")
             and type(process.get("InteractionCount")) is int and process["InteractionCount"] == len(interactions)
             and type(process.get("QueuedReplyCount")) is int and process["QueuedReplyCount"] == len(replies)
             and type(process.get("ReplyCount")) is int and process["ReplyCount"] == len(replies), "Interactive writer/events/counters incomplete")
        need([event.get("sequence") for event in interactions] == list(range(1, len(interactions) + 1))
             and all(event.get("protocol") == "winbooksplit.interaction" and type(event.get("version")) is int and event["version"] == 1
                     and isinstance(event.get("session"), str) and re.fullmatch(r"[0-9a-f]{32}", event["session"]) for event in interactions),
             "Captured interaction frames have invalid sequence/protocol/session")
        for index, reply in enumerate(replies):
            need(index < len(interactions) and all(reply.get(key) == interactions[index].get(key) for key in ("protocol", "version", "session", "sequence")),
                 "Actual reply did not bind to its expected interaction")
        if kind == "fallback-manual":
            need([event.get("stage") for event in interactions] == ["no_plan", "input_ready", "plan_ready"]
                 and interactions[1].get("source_identity") == interactions[2]["plan"].get("source_identity")
                 and replies[0].get("action") == "retry" and replies[0].get("mode") == "manual", "Fallback did not explicitly reuse its captured source")
    if kind == "pending-timeout":
        need(final.get("status") == "timeout" and final.get("code") == "processor_timeout" and process.get("TimedOut") is True
             and final.get("engine_result") is None and len(rows) <= 1 and "-NoPause" not in parameters
             and "Press Enter to exit" not in row["stdout"] and row.get("stdin_utf8") == ""
             and replies == [] and process.get("QueuedReplyCount") == process.get("ReplyCount") == 0
             and type(process.get("ResultRecordCount")) is int and process["ResultRecordCount"] == len(rows)
             and type(process.get("ExitCode")) is int and process["ExitCode"] != 0,
             "Pending-input timeout did not terminate cleanly without pause")
        if rows:
            # Closing owned stdin during timeout can let EOF cancellation emit
            # one frame before forced shutdown wins the native-exit race.
            frame = rows[0]
            need(frame.get("protocol") == "winbooksplit.result" and type(frame.get("version")) is int and frame["version"] == 1
                 and frame.get("mode") == final.get("mode") == "2" and frame.get("status") == "cancelled"
                 and frame.get("code") == "processing_cancelled" and type(frame.get("exit_code")) is int and frame["exit_code"] == 130
                 and type(frame.get("written_count")) is int and frame["written_count"] == 0 and frame.get("execution") is None
                 and frame.get("diagnostic") is None and frame.get("fallback_modes") == [] and frame.get("warnings") == [],
                 "Timeout captured success, execution or a different terminal frame")
    else:
        if kind != "preview":
            need(process.get("TimedOut") is False and process.get("Cancelled") is False and len(rows) == 1
                 and rows[0] == final.get("engine_result"), "Terminal result/native transport evidence differs")
        frame = rows[0]
        need(frame.get("protocol") == "winbooksplit.result" and type(frame.get("version")) is int and frame["version"] == 1
             and type(frame.get("exit_code")) is int and frame["exit_code"] == (process.get("ExitCode") if process is not None else row["exit_code"])
             and type(frame.get("written_count")) is int and frame.get("mode") == final.get("mode"), "Terminal engine protocol/native status differs")
    if unavailable:
        need(row.get("hold") is None and plans == replies == [] and row.get("stdin_utf8") == ""
             and not any(token in row["stdout"] for token in (PROMPT, "Open the generated PDF?", "Open output folder?", "Press Enter to exit")),
             "NonInteractive/Preview prompted or opened/pause flow")
    elif kind != "fallback-cancel":
        hold = row.get("hold")
        need(isinstance(hold, dict) and hold.get("parent_running") is True and hold.get("chapter_or_stage_absent") is True
             and isinstance(hold.get("stdout"), str) and PROMPT in hold["stdout"] and "Done." not in hold["stdout"]
             and hold.get("console_members") == [".WinBookSplit-console-owner.json", "console.log"], "No-write held confirmation proof missing")
        console = [name for name in hold.get("base_members", []) if name.startswith(".WinBookSplit-console-")]
        need(len(console) == 1 and set(hold["base_members"]) == {"prior-output.pdf", console[0]}
             and hold.get("console_owner") == {"run_id": console[0].removeprefix(".WinBookSplit-console-"), "kind": "console"},
             "Held base contains unapproved chapter/stage or ownership mismatch")
    if kind == "fallback-cancel":
        need(final.get("status") == "cancelled" and final.get("code") == "cancelled" and rows[0].get("status") == "no_plan"
             and rows[0].get("exit_code") == process.get("ExitCode") == 5 and plans == []
             and [reply.get("action") for reply in replies] == ["cancel"], "Fallback cancellation lost its original failed attempt")
    elif kind in CANCELLED:
        need(len(plans) == 1 and final.get("written_count") == 0 and final.get("final_directory") is None
             and row.get("publication") is None and "Done." not in row["stdout"], "Cancelled plan published or announced success")
        if kind != "pending-timeout":
            need(rows[0].get("status") == "cancelled" and rows[0].get("exit_code") == 130
                 and replies[-1].get("action") == "cancel", "Confirmation cancellation was not preserved")
    elif kind == "preview":
        need(final.get("status") == "preview" and final.get("code") == rows[0].get("code") == "preview_complete"
             and rows[0].get("status") == "preview" and rows[0].get("execution") is None and rows[0].get("written_count") == 0
             and final["written_count"] == 0 and final.get("final_directory") is None and rows[0].get("fallback_modes") == []
             and row.get("publication") is None and "Done." not in row["stdout"] and "Output: " not in row["stdout"]
             and "Preview only: no chapter PDFs were written." in row["stdout"], "Preview wrote or announced execution success")
        plan = rows[0].get("plan")
        cli.coverage(plan, NESTED_RANGES)
        identity = plan.get("source_identity", {})
        observed_source = row["source_observations_before"]["source"]
        titles, reasons = expected_entries(kind)
        need(plan.get("mode") == final.get("mode") == "2" and plan.get("output_naming", {}).get("resolved_base") == row["output_base"]
             and identity.get("binding") == "reader_snapshot" and identity.get("path") == identity.get("resolved_path") == row["input_path"]
             and identity.get("sha256") == observed_source["sha256"] and identity.get("size_bytes") == observed_source["size_bytes"]
             and [entry.get("title") for entry in plan["entries"]] == titles and [entry.get("reason") for entry in plan["entries"]] == reasons,
             "Preview plan differs from actual immutable source/destination/nested semantics")
        need(f"Physical PDF pages: {plan['total_pages']}" in row["stdout"] and "Output base: " + row["output_base"] in row["stdout"]
             and f"Coverage: every physical page exactly once; {plan['total_pages']} pages, {len(plan['entries'])} sections." in row["stdout"],
             "Preview human physical count/destination/coverage differs")
        for entry in plan["entries"]:
            name = entry.get("filename")
            need(isinstance(name, str) and name not in {".pdf", "..pdf"} and name == f"{entry['sequence']:02d} - {entry['title']}.pdf"
                 and name == name.strip() and not re.search(r'[<>:"/\\|?*\x00-\x1f]', name)
                 and not any(unicodedata.category(char) in {"Cc", "Cf", "Cs"} for char in name)
                 and f"{entry['sequence']}. {entry['title']} | physical pages {entry['start'] + 1}-{entry['end']} | {name} | {entry['reason']}" in row["stdout"],
                 "Preview human title/range/filename/reason differs or basename is unsafe")
    else:
        need(final.get("status") == "success" and final.get("code") == rows[0].get("code") == "split_complete"
             and rows[0].get("status") == "success" and final.get("written_count", 0) > 0 and "Done." in row["stdout"], "Actual UX execution failed or zero outputs")
        mode = "manual" if kind in {"manual-physical", "ebook-manual-epub", "ebook-manual-azw3", "fallback-manual"} else "2"
        execution = rows[0].get("execution")
        plan = validate_plan_event(plans[0], mode, row["output_base"]) if not unavailable else execution
        ranges = expected_ranges(kind)
        if not unavailable:
            need(len(plans) == 1 and replies[-1].get("action") == "execute" and replies[-1].get("plan_sha256") == plans[0]["plan_sha256"],
                 "Output lacked explicit consent for the exact displayed plan")
            cli.coverage(plan, ranges)
            titles, reasons = expected_entries(kind)
            need([entry.get("title") for entry in plan["entries"]] == titles and [entry.get("reason") for entry in plan["entries"]] == reasons,
                 "Displayed nested/manual semantics differ from independent authored oracle")
            for planned, written in zip(plan["entries"], execution["outputs"]):
                need(all(planned.get(key) == written.get(key) for key in ("sequence", "title", "filename", "start", "end", "parent_id", "reason", "warnings")),
                     "Writer changed confirmed title/filename/boundary metadata")
            need(plan["source_identity"] == execution.get("source_identity"), "Writer changed captured reader identity")
        need(execution.get("mode") == mode and execution.get("total_pages") == ranges[-1][1]
             and execution.get("coverage") == {"complete": True, "covered_pages": ranges[-1][1], "section_count": len(ranges)}
             and [[entry.get("start"), entry.get("end")] for entry in execution.get("outputs", [])] == ranges
             and execution.get("written_count") == final["written_count"] == len(ranges), "Executed coverage/count differs")
        published = row.get("publication")
        need(isinstance(published, dict) and published.get("manifest") == execution.get("manifest")
             and isinstance(published.get("manifest_raw"), str) and json.loads(published["manifest_raw"]) == execution["manifest"]
             and published.get("manifest_sha256") == sha256(published["manifest_raw"].encode("utf-8")).hexdigest()
             and published.get("content_sha256") == row.get("expected_content_sha256")
             and len(published["content_sha256"]) == execution["total_pages"], "Actual reopened chapter content/manifest proof missing")
        need(all(execution.get(key) == execution["manifest"].get(key) for key in
                 ("run_id", "final_directory", "mode", "total_pages", "source_identity", "coverage", "written_count", "outputs"))
             and execution["manifest"].get("status") == "complete"
             and published.get("members") == sorted([".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", *[entry["filename"] for entry in execution["outputs"]]]),
             "Actual manifest/member fields contradict execution")
        need(f"Chapters written: {len(ranges)}; physical pages: {ranges[-1][1]}; every page exactly once." in row["stdout"],
             "Truthful final physical count summary missing")
    if plans:
        mode = plans[0]["mode"]
        plan = validate_plan_event(plans[0], mode, row["output_base"])
        cli.coverage(plan, expected_ranges(kind))
        hold = row.get("hold")
        need(isinstance(hold, dict) and f"Physical PDF pages: {plan['total_pages']}" in hold["stdout"]
             and "Output base: " + row["output_base"] in hold["stdout"]
             and f"Coverage: every physical page exactly once; {plan['total_pages']} pages, {len(plan['entries'])} sections." in hold["stdout"],
             "Human plan lacked physical count/destination/coverage before consent")
        for entry in plan["entries"]:
            need(f"{entry['sequence']}. {entry['title']} | physical pages {entry['start'] + 1}-{entry['end']} | {entry['filename']} | {entry['reason']}" in hold["stdout"],
                 "Human plan title/range/filename/reason differed before consent")
    if kind == "confirm-invalid":
        need(row["stdout"].count("Choose exactly Y or C.") == 2 and row["stdin_utf8"] == "Q\nmaybe\nY\nN\n", "Invalid consent was silently accepted or altered")
    if kind.startswith("ebook-"):
        need("Use these numbers, not ebook locations." in row["stdout"] and "Open the generated PDF?" in row["stdout"]
             and row["stdin_utf8"] == "M\nN\n1,2,3\nY\nN\n"
             and not any(reply.get("action") == "request_working_pdf" for reply in replies), "Ebook numbering/explicit declined opening differs")


def validate_working_case(row):
    need(isinstance(row, dict) and row.get("format") in {"epub", "azw3"} and row.get("id") == "engine-working-" + row["format"], "Unknown working PDF case")
    raw_streams(row)
    need(row.get("viewer_opened") is False and all(row.get(key) is True for key in ("working_readable_before_starts",
         "working_identity_unchanged_at_plan", "working_pdf_absent_after", "working_directory_absent_after"))
         and row.get("source_before") == row.get("source_after") and row.get("output_members_after") == ["prior-output.pdf"],
         "Working PDF lifetime/cleanup/source proof missing or viewer claim fabricated")
    events = [json.loads(line[len("[WBS-INTERACTION] "):]) for line in row["stdout"].splitlines() if line.startswith("[WBS-INTERACTION] ")]
    need(events == row.get("events") and [event.get("stage") for event in events] == ["input_ready", "working_pdf_ready", "plan_ready"]
         and [event.get("sequence") for event in events] == [1, 2, 3]
         and all(event.get("session") == "b" * 32 and event.get("protocol") == "winbooksplit.interaction" and event.get("version") == 1 for event in events),
         "Working PDF event sequence/protocol differs")
    working, identity = events[1]["working_pdf"], row.get("working_identity")
    need(working.get("binding") == "working_pdf_copy" and isinstance(identity, dict) and identity.get("sha256") == working.get("sha256")
         == events[0]["source_identity"]["sha256"] and identity.get("size_bytes") == working.get("size_bytes")
         and row.get("working_content_sha256") == row.get("reference_content_sha256")
         and len(row["working_content_sha256"]) == working.get("page_count") == (3 if row["format"] == "epub" else 4),
         "Actual working copy did not match captured/reference physical content")
    need(events[0].get("conversion", {}).get("workspace_cleanup") == {"cleanup_complete": True, "retained_staging": None},
         "Original conversion cleanup was not completed before working copy")
    result = row.get("terminal_result")
    need([json.loads(line) for line in row["stdout"].splitlines() if line.startswith("{")] == [result]
         and result.get("protocol") == "winbooksplit.result" and result.get("status") == "cancelled"
         and result.get("exit_code") == row["exit_code"] == 130 and result.get("execution") is None
         and result.get("written_count") == 0 and result.get("diagnostic", {}).get("working_pdf_cleanup", {}).get("cleanup_complete") is True,
         "Working PDF cancellation/cleanup terminal result differs")
    validate_plan_event(events[-1], "manual", events[-1]["plan"]["output_naming"]["resolved_base"])


def validate_ux_report(report, requested_shells, calibre):
    need(isinstance(report, dict) and report.get("schema_version") == 1 and report.get("task_id") == "M3-T04"
         and report.get("result") == "UX_REGRESSION_PASSED" and report.get("success") is True and report.get("exit_code") == 0
         and report.get("acceptance_ids") == ["AC-060", "AC-061", "AC-062"] and report.get("source_unchanged") is True
         and report.get("owned_temp_removed") is True and report.get("human_or_viewer_opening_tested") is False,
         "UX top-level pass/source/cleanup/scope proof missing")
    manifest = report.get("tested_path_sha256")
    need(cli.hashes(manifest) and APPLICATION | {"tests/ux/characterize_ux.py", "tests/ux/validate_ux_report.py", "tests/run_tests.py"} <= set(manifest),
         "UX raw source/validator/runner hashes missing")
    hosts = cli.exact_cases(report.get("host_cases"), {"PS51", "PS7"})
    need(len(requested_shells) == 2 and len(set(requested_shells)) == 2 and {host.get("shell_executable") for host in hosts.values()} == set(requested_shells),
         "Requested native UX hosts differ")
    for name, host in hosts.items():
        need(host.get("passed") is True and host.get("exit_code") == 0 and host.get("host_major") == (5 if name == "PS51" else 7)
             and host.get("syntax_error_count") == 0 and host.get("stored_policies") == host.get("policies_after"), "Host syntax/version/policies differ")
    cases = cli.exact_cases(report.get("cases"), {host + "-" + kind for host in hosts for kind in APPLICATION_KINDS})
    for row in cases.values():
        validate_application_case(row)
        host = hosts[row["host_id"]]
        need(row["host_version"] == host["host_version"] and row["shell_executable"] == host["shell_executable"]
             and all(manifest[name] == digest for name, digest in row["application_sha256"].items()), "UX case source/requested host differs")
    working = cli.exact_cases(report.get("working_pdf_cases"), {"engine-working-epub", "engine-working-azw3"})
    for row in working.values():
        validate_working_case(row)
        need("--calibre-path" in row["command"] and row["command"][row["command"].index("--calibre-path") + 1] == calibre,
             "Working copy did not use requested real converter")
    need(report.get("case_count") == len(cases) + len(working), "UX native case count differs")
    fixtures = report.get("fixture_provenance", {})
    need(fixtures.get("authored_original") is True and fixtures.get("remote_resources") is False
         and fixtures.get("calibre", {}).get("path") == calibre and fixtures["calibre"].get("version") == "9.15.0"
         and fixtures.get("azw3_generation", {}).get("exit_code") == 0, "Real offline ebook fixture provenance missing")
    references = report.get("real_conversion_references")
    need(isinstance(references, dict) and set(references) == {"epub", "azw3"}, "Independent real conversion references missing")
    for fmt, reference in references.items():
        need(reference.get("process", {}).get("exit_code") == 0 and reference["process"].get("argv", [None])[0] == calibre
             and reference.get("page_count") == (3 if fmt == "epub" else 4)
             and HEX.fullmatch(reference.get("sha256", "")) and len(reference.get("page_content_sha256", [])) == reference["page_count"],
             "Independent real converter reference invalid")
        need(working["engine-working-" + fmt]["reference_content_sha256"] == reference["page_content_sha256"],
             "Working PDF reference content was substituted")
        for row in cases.values():
            if row["kind"] == "ebook-manual-" + fmt:
                need(row["expected_content_sha256"] == reference["page_content_sha256"], "Application ebook content reference was substituted")
