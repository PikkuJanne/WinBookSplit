"""Synthetic mutation controls only; these do not assert actual native acceptance."""

from copy import deepcopy
from hashlib import sha256
import importlib.util
import json
from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("wbs_ux_unit_validator", ROOT / "tests/ux/validate_ux_report.py")
validator = importlib.util.module_from_spec(spec)
spec.loader.exec_module(validator)
spec = importlib.util.spec_from_file_location("wbs_ux_console_fixture", ROOT / "tests/python/console_receipt_fixture.py")
console_fixture = importlib.util.module_from_spec(spec)
spec.loader.exec_module(console_fixture)


def compact(value):
    return json.dumps(value, ensure_ascii=True, separators=(",", ":"))


def synthetic_cancel():
    source, base = "C:/synthetic/source.pdf", "C:/synthetic/out"
    titles, reasons = validator.expected_entries("confirm-cancel")
    entries = [{"sequence": number, "title": title, "reason": reason, "start": start, "end": end,
                "filename": f"{number:02d} - {title}.pdf", "parent_id": None, "warnings": []}
               for number, ((start, end), title, reason) in enumerate(zip(validator.NESTED_RANGES, titles, reasons), 1)]
    plan = {"mode": "2", "total_pages": 12, "entries": entries, "ranges": deepcopy(validator.NESTED_RANGES),
            "source_identity": {"path": source, "resolved_path": source, "sha256": "a" * 64, "size_bytes": 10, "binding": "reader_snapshot"},
            "output_naming": {"resolved_base": base}, "coverage": {"complete": True, "covered_pages": 12, "section_count": 5}}
    canonical = json.dumps(plan, sort_keys=True, ensure_ascii=True, separators=(",", ":"))
    event = {"protocol": "winbooksplit.interaction", "version": 1, "session": "a" * 32,
             "sequence": 1, "stage": "plan_ready", "mode": "2", "plan": plan,
             "plan_json": canonical, "plan_sha256": sha256(canonical.encode()).hexdigest()}
    reply = {"protocol": "winbooksplit.interaction", "version": 1, "session": "a" * 32, "sequence": 1, "action": "cancel"}
    result = {"protocol": "winbooksplit.result", "version": 1, "mode": "2", "status": "cancelled",
              "code": "processing_cancelled", "exit_code": 130, "written_count": 0, "execution": None,
              "diagnostic": None, "warnings": [], "fallback_modes": []}
    final = {"protocol": "winbooksplit.outcome", "version": 1, "mode": "2", "status": "cancelled", "code": "processing_cancelled",
             "exit_code": 130, "written_count": 0, "final_directory": None, "engine_result": result}
    process = {"JobAssigned": True, "ParentStopped": True, "DescendantsStopped": True, "StreamsComplete": True,
               "TimedOut": False, "Cancelled": False, "ExitCode": 130, "InputWriterStopped": True,
               "InteractionCount": 1, "QueuedReplyCount": 1, "ReplyCount": 1}
    observations = {name: {"sha256": "a" * 64, "size_bytes": 10, "device": 1, "inode": number, "attributes": 1}
                    for number, name in enumerate(("source", "neighbor", "prior"), 1)}
    parameters = ["-InputFile", source, "-OutputDirectory", base, "-PythonPath", "C:/synthetic/python.exe",
                  "-Mode", "Auto", "-BookmarkLevel", "2", "-NoPause"]
    display = "Physical PDF pages: 12\nOutput base: " + base + "\n"
    display += "\n".join(f"{entry['sequence']}. {entry['title']} | physical pages {entry['start'] + 1}-{entry['end']} | {entry['filename']} | {entry['reason']}" for entry in entries)
    display += "\nCoverage: every physical page exactly once; 12 pages, 5 sections.\n" + validator.PROMPT
    case = {"id": "PS51-confirm-cancel", "kind": "confirm-cancel", "host_id": "PS51", "host_version": "5.1.26100.9444",
            "shell_executable": "C:/synthetic/powershell.exe", "parameters": parameters,
            "command": ["C:/synthetic/powershell.exe", "-NoProfile", "-File", "C:/synthetic/WinBookSplit.ps1", *parameters],
            "cwd": "C:/synthetic/unrelated", "pid": 1234, "actual_process": True, "exit_code": 130, "timed_out": False,
            "streams_complete": True, "elapsed_seconds": 1, "timeout_seconds": 90, "stdin_utf8": "C\n",
            "passed": True, "application_unchanged": True, "source_read_only_observed": True,
            "application_sha256": {name: "b" * 64 for name in validator.APPLICATION},
            "input_path": source, "output_base": base, "source_observations_before": observations,
            "source_observations_after": deepcopy(observations), "source_observations_after_cleanup": deepcopy(observations),
            "output_members_after_cleanup": ["prior-output.pdf"], "hold": {"parent_running": True,
                "chapter_or_stage_absent": True, "stdout": display,
                "console_members": [".WinBookSplit-console-owner.json", "console.log"],
                "base_members": [".WinBookSplit-console-" + "c" * 32, "prior-output.pdf"],
                "console_owner": {"run_id": "c" * 32, "kind": "console"}},
            "outcome": final, "engine_records": [result], "process_summaries": [process],
            "plan_events": [event], "interaction_replies": [reply], "interaction_events": [deepcopy(event)],
            "publication": None, "expected_content_sha256": ["d" * 64] * 12, "stderr": ""}
    reseal(case)
    return case


def reseal(case):
    lines = ["[PLAN] " + compact(event) for event in case["plan_events"]]
    lines += ["[INTERACTION-REPLY] " + compact(reply) for reply in case["interaction_replies"]]
    lines += ["[WBS-INTERACTION] " + compact(event) for event in case["interaction_events"]]
    lines += [compact(result) for result in case["engine_records"]]
    lines += ["[PROCESS] " + compact(process) for process in case["process_summaries"]]
    lines += ["[OPERATION-OUTCOME] " + compact(case["outcome"])]
    case["console_log"] = "\r\n".join(lines) + "\r\n"
    case["console_log_sha256"] = sha256(case["console_log"].encode()).hexdigest()
    case["console_evidence"] = console_fixture.make(case["outcome"], case["console_log_sha256"],
        log_text=case["console_log"], plan=case["plan_events"][-1]["plan"] if case["plan_events"] else None,
        source_path=case["input_path"], source_observation=case["source_observations_before"]["source"])
    case["stdout"] = case.get("stdout_prefix", "") + "[OUTCOME] " + compact(case["outcome"]) + "\r\n"
    for name in ("stdout", "stderr"):
        case[name + "_sha256"] = sha256(case[name].encode()).hexdigest()



def canonical(event):
    event["plan_json"] = json.dumps(event["plan"], sort_keys=True, ensure_ascii=True, separators=(",", ":"))
    event["plan_sha256"] = sha256(event["plan_json"].encode()).hexdigest()


def synthetic_success():
    case = synthetic_cancel()
    event = case["plan_events"][0]
    plan = event["plan"]
    folder = case["output_base"] + "/published"
    outputs = deepcopy(plan["entries"])
    manifest = {"run_id": "e" * 32, "final_directory": folder, "mode": "2", "total_pages": 12,
                "source_identity": deepcopy(plan["source_identity"]), "coverage": deepcopy(plan["coverage"]),
                "written_count": 5, "outputs": outputs, "status": "complete"}
    execution = {**deepcopy(manifest), "manifest": deepcopy(manifest)}
    result = case["engine_records"][0]
    result.update(status="success", code="split_complete", exit_code=0, written_count=5, execution=execution)
    case["outcome"].update(status="success", code="split_complete", exit_code=0, written_count=5, final_directory=folder)
    case.update(id="PS51-nested-confirm", kind="nested-confirm", exit_code=0, stdin_utf8="Y\nN\n")
    case["process_summaries"][0]["ExitCode"] = 0
    case["interaction_replies"][0].update(action="execute", plan_sha256=event["plan_sha256"])
    raw = compact(manifest)
    case["publication"] = {"manifest": deepcopy(manifest), "manifest_raw": raw,
                           "manifest_sha256": sha256(raw.encode()).hexdigest(),
                           "members": sorted([".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", *[entry["filename"] for entry in outputs]]),
                           "content_sha256": case["expected_content_sha256"]}
    case["stdout_prefix"] = "Done.\nOutput: " + folder + "\nChapters written: 5; physical pages: 12; every page exactly once.\n"
    reseal(case)
    return case


def synthetic_noninteractive():
    case = synthetic_success()
    case.update(id="PS51-noninteractive", kind="noninteractive", stdin_utf8="", hold=None,
                plan_events=[], interaction_events=[], interaction_replies=[])
    case["parameters"].append("-NonInteractive")
    case["command"] = [case["shell_executable"], "-NoProfile", "-NonInteractive", "-File",
                       "C:/synthetic/WinBookSplit.ps1", *case["parameters"]]
    case["process_summaries"][0].update(InteractionCount=0, QueuedReplyCount=0, ReplyCount=0)
    reseal(case)
    return case


def synthetic_fallback():
    case = synthetic_cancel()
    result = case["engine_records"][0]
    result.update(mode="1", status="no_plan", code="no_bookmarks", exit_code=5, fallback_modes=["manual"])
    case["outcome"].update(mode="1", code="cancelled")
    case.update(id="PS51-fallback-cancel", kind="fallback-cancel", hold=None, plan_events=[])
    case["process_summaries"][0]["ExitCode"] = 5
    case["interaction_events"] = [{"protocol": "winbooksplit.interaction", "version": 1, "session": "a" * 32,
                                   "sequence": 1, "stage": "no_plan", "mode": "1", "result": deepcopy(result), "fallback_modes": ["manual"]}]
    reseal(case)
    return case


def reseal_working(case):
    case["stdout"] = "".join("[WBS-INTERACTION] " + compact(event) + "\n" for event in case["events"])
    case["stdout"] += compact(case["terminal_result"]) + "\n"
    for name in ("stdout", "stderr"):
        case[name + "_sha256"] = sha256(case[name].encode()).hexdigest()


def synthetic_working(fmt):
    original_path, base = "C:/synthetic/source." + fmt, "C:/synthetic/out"
    captured = {"path": base + "/converted.pdf", "resolved_path": base + "/converted.pdf",
                "binding": "reader_snapshot", "sha256": "a" * 64, "size_bytes": 20}
    original = {"path": original_path, "resolved_path": original_path, "binding": "ebook_snapshot", "sha256": "b" * 64, "size_bytes": 10}
    ranges = validator.expected_ranges("ebook-manual-" + fmt)
    titles, reasons = validator.expected_entries("ebook-manual-" + fmt)
    entries = [{"sequence": number, "title": title, "reason": reason, "start": start, "end": end,
                "filename": f"{number:02d} - {title}.pdf", "parent_id": None, "warnings": []}
               for number, ((start, end), title, reason) in enumerate(zip(ranges, titles, reasons), 1)]
    plan = {"mode": "manual", "total_pages": ranges[-1][1], "entries": entries, "ranges": ranges,
            "source_identity": deepcopy(captured), "original_ebook_identity": deepcopy(original), "output_naming": {"resolved_base": base},
            "coverage": {"complete": True, "covered_pages": ranges[-1][1], "section_count": 3}}
    header = {"protocol": "winbooksplit.interaction", "version": 1, "session": "b" * 32, "mode": "manual"}
    working = {"path": base + "/working.pdf", "binding": "working_pdf_copy", "sha256": "a" * 64,
               "size_bytes": 20, "page_count": ranges[-1][1]}
    events = [{**header, "sequence": 1, "stage": "input_ready", "source_identity": deepcopy(captured),
               "original_ebook_identity": deepcopy(original), "conversion": {"workspace_cleanup": {"cleanup_complete": True, "retained_staging": None}}},
              {**header, "sequence": 2, "stage": "working_pdf_ready", "source_identity": deepcopy(captured), "working_pdf": working},
              {**header, "sequence": 3, "stage": "plan_ready", "plan": plan}]
    canonical(events[-1])
    result = {"protocol": "winbooksplit.result", "version": 1, "mode": "manual", "status": "cancelled",
              "code": "processing_cancelled", "exit_code": 130, "written_count": 0, "execution": None,
              "diagnostic": {"working_pdf": deepcopy(working), "working_pdf_cleanup": {"cleanup_complete": True}}}
    observed = {"sha256": "b" * 64, "size_bytes": 10, "device": 1, "inode": 1, "attributes": 32}
    case = {"id": "engine-working-" + fmt, "format": fmt, "viewer_opened": False,
            "command": ["C:/synthetic/python.exe", "C:/synthetic/winbooksplit_engine.py", original_path, base, "manual", "", "--interactive", "b" * 32],
            "events": events, "terminal_result": result, "working_identity": {"sha256": "a" * 64, "size_bytes": 20},
            "working_content_sha256": ["c" * 64] * ranges[-1][1], "reference_content_sha256": ["c" * 64] * ranges[-1][1],
            "working_readable_before_starts": True, "working_identity_unchanged_at_plan": True,
            "working_pdf_absent_after": True, "working_directory_absent_after": True,
            "source_before": observed, "source_after": deepcopy(observed), "output_members_after": ["prior-output.pdf"],
            "actual_process": True, "pid": 1234, "exit_code": 130, "timed_out": False, "streams_complete": True,
            "elapsed_seconds": 1, "timeout_seconds": 90, "cwd": "C:/synthetic", "stderr": ""}
    reseal_working(case)
    return case


class UxReceiptTests(unittest.TestCase):
    def test_synthetic_cancel_receipt_requires_actual_held_no_write_and_native_identity_proofs(self):
        case = synthetic_cancel()
        validator.validate_application_case(case)
        mutations = [lambda row: row.update(exit_code=0), lambda row: row.update(stdin_utf8="Y\n"),
                     lambda row: row.update(timed_out=True), lambda row: row.update(pid=True),
                     lambda row: row.update(streams_complete=False), lambda row: row.update(application_unchanged=False),
                     lambda row: row.update(source_observations_after={}),
                     lambda row: row["hold"].update(parent_running=False),
                     lambda row: row["hold"].update(base_members=["prior-output.pdf", "chapter.pdf"]),
                     lambda row: row["hold"].update(stdout=validator.PROMPT),
                     lambda row: row["process_summaries"][0].update(DescendantsStopped=False),
                     lambda row: row["process_summaries"][0].update(InputWriterStopped=False),
                     lambda row: row["process_summaries"][0].update(InteractionCount=2),
                     lambda row: row["interaction_replies"][0].update(session="b" * 32)]
        for index, mutate in enumerate(mutations):
            changed = deepcopy(case)
            mutate(changed)
            reseal(changed)
            with self.subTest(index=index), self.assertRaises(ValueError):
                validator.validate_application_case(changed)

    def test_plan_digest_partition_and_human_text_cannot_be_forged_by_resealing_outer_streams(self):
        case = synthetic_cancel()
        for mutate in (lambda row: row["plan_events"][0].update(plan_sha256="0" * 64),
                       lambda row: row["plan_events"][0]["plan"]["entries"][1].update(start=4),
                       lambda row: row["plan_events"][0]["plan"]["coverage"].update(covered_pages=11),
                       lambda row: row["hold"].update(stdout=row["hold"]["stdout"].replace("physical pages 3-3", "physical pages 4-4")),
                       lambda row: row["interaction_events"][0].update(sequence=2)):
            changed = deepcopy(case)
            mutate(changed)
            reseal(changed)
            with self.assertRaises(ValueError):
                validator.validate_application_case(changed)
        for mutate in (lambda plan: plan["entries"][1].update(start=4),
                       lambda plan: plan["coverage"].update(covered_pages=11),
                       lambda plan: plan["entries"][0].update(filename="../foreign.pdf")):
            changed = deepcopy(case)
            event = changed["plan_events"][0]
            mutate(event["plan"])
            event["plan_json"] = json.dumps(event["plan"], sort_keys=True, ensure_ascii=True, separators=(",", ":"))
            event["plan_sha256"] = sha256(event["plan_json"].encode()).hexdigest()
            changed["interaction_events"] = [deepcopy(event)]
            reseal(changed)
            with self.assertRaises(ValueError):
                validator.validate_application_case(changed)

    def test_final_confirmation_cancellation_status_and_zero_output_cannot_be_relabelled(self):
        for kind, answer in (("confirm-cancel", "C\n"), ("confirm-blank", "\n"), ("confirm-eof", "")):
            case = synthetic_cancel()
            case.update(kind=kind, id="PS51-" + kind, stdin_utf8=answer)
            validator.validate_application_case(case)
            mutations = [lambda row: row["outcome"].update(status="success"),
                         lambda row: row["outcome"].update(status="failed"),
                         lambda row: row["outcome"].update(status="timeout"),
                         lambda row: row["engine_records"][0].update(written_count=1),
                         lambda row: row["engine_records"][0].update(execution={"written_count": 1, "outputs": ["fabricated"]}),
                         lambda row: (row["outcome"].update(mode="manual"), row["engine_records"][0].update(mode="manual"))]
            for index, mutation in enumerate(mutations):
                changed = deepcopy(case)
                mutation(changed)
                reseal(changed)
                with self.subTest(kind=kind, mutation=index), self.assertRaises(ValueError):
                    validator.validate_application_case(changed)

    def test_final_and_engine_confirmation_cancellation_codes_cannot_claim_success(self):
        for kind, answer in (("confirm-cancel", "C\n"), ("confirm-blank", "\n"), ("confirm-eof", "")):
            case = synthetic_cancel()
            case.update(kind=kind, id="PS51-" + kind, stdin_utf8=answer)
            validator.validate_application_case(case)
            for scope in ("outcome", "engine_records"):
                for code in ("split_complete", "preview_complete", "cancelled", "processor_cancelled", "no_bookmarks"):
                    changed = deepcopy(case)
                    target = changed["outcome"] if scope == "outcome" else changed["engine_records"][0]
                    target["code"] = code
                    reseal(changed)
                    with self.subTest(kind=kind, scope=scope, code=code), self.assertRaises(ValueError):
                        validator.validate_application_case(changed)

    def test_success_terminal_count_and_mode_match_the_held_execution(self):
        case = synthetic_success()
        validator.validate_application_case(case)
        for mutation in ("count", "mode"):
            changed = deepcopy(case)
            if mutation == "count":
                changed["engine_records"][0]["written_count"] = 1
            else:
                changed["engine_records"][0]["mode"] = "manual"
                changed["outcome"]["mode"] = "manual"
            reseal(changed)
            with self.subTest(mutation=mutation), self.assertRaisesRegex(ValueError, "mode/coverage/count"):
                validator.validate_application_case(changed)

    def test_noninteractive_executed_source_matches_the_actual_input_observation(self):
        case = synthetic_noninteractive()
        validator.validate_application_case(case)
        for field, value in (("sha256", "f" * 64), ("size_bytes", 11), ("path", "C:/synthetic/foreign.pdf")):
            changed = deepcopy(case)
            execution = changed["engine_records"][0]["execution"]
            execution["source_identity"][field] = value
            execution["manifest"]["source_identity"][field] = value
            publication = changed["publication"]
            publication["manifest"]["source_identity"][field] = value
            publication["manifest_raw"] = compact(publication["manifest"])
            publication["manifest_sha256"] = sha256(publication["manifest_raw"].encode()).hexdigest()
            reseal(changed)
            with self.subTest(field=field), self.assertRaisesRegex(ValueError, "actual immutable source"):
                validator.validate_application_case(changed)

    def test_consistently_resealed_publication_cannot_leave_the_confirmed_output_base(self):
        case = synthetic_success()
        validator.validate_application_case(case)
        changed = deepcopy(case)
        old = changed["outcome"]["final_directory"]
        folder = "C:/synthetic/foreign/published"
        execution = changed["engine_records"][0]["execution"]
        changed["outcome"]["final_directory"] = folder
        execution["final_directory"] = execution["manifest"]["final_directory"] = folder
        publication = changed["publication"]
        publication["manifest"]["final_directory"] = folder
        publication["manifest_raw"] = compact(publication["manifest"])
        publication["manifest_sha256"] = sha256(publication["manifest_raw"].encode()).hexdigest()
        changed["stdout_prefix"] = changed["stdout_prefix"].replace(old, folder)
        reseal(changed)
        with self.assertRaisesRegex(ValueError, "output directory"):
            validator.validate_application_case(changed)

    def test_success_final_and_displayed_output_directory_bind_to_execution(self):
        case = synthetic_success()
        validator.validate_application_case(case)
        for mutate in (lambda row: row["outcome"].update(final_directory="C:/synthetic/foreign"),
                       lambda row: row.update(stdout_prefix=row["stdout_prefix"].replace("/published", "/foreign")),
                       lambda row: row.update(stdout_prefix=row["stdout_prefix"] + "Output: C:/synthetic/foreign\n")):
            changed = deepcopy(case)
            mutate(changed)
            reseal(changed)
            with self.assertRaisesRegex(ValueError, "output directory"):
                validator.validate_application_case(changed)

    def test_fallback_cancel_preserves_the_exact_no_plan_attempt_without_output(self):
        case = synthetic_fallback()
        validator.validate_application_case(case)
        for mutate in (lambda row: row["engine_records"][0].update(code="split_complete"),
                       lambda row: row["interaction_events"][0]["result"].update(code="split_complete"),
                       lambda row: row["interaction_events"][0]["result"].update(message="different attempt"),
                       lambda row: row["engine_records"][0].update(written_count=1),
                       lambda row: row["engine_records"][0].update(execution={"written_count": 1}),
                       lambda row: row["outcome"].update(final_directory="C:/synthetic/foreign"),
                       lambda row: row.update(stdout_prefix="Done.\n")):
            changed = deepcopy(case)
            mutate(changed)
            reseal(changed)
            with self.assertRaisesRegex(ValueError, "original failed attempt"):
                validator.validate_application_case(changed)

    def test_logged_plan_and_captured_source_remain_bound_after_resealing(self):
        case = synthetic_cancel()
        validator.validate_application_case(case)
        for change_raw_event in (False, True):
            changed = deepcopy(case)
            changed["plan_events"][0]["plan"]["source_identity"]["sha256"] = "f" * 64
            canonical(changed["plan_events"][0])
            if change_raw_event:
                changed["interaction_events"] = deepcopy(changed["plan_events"])
            reseal(changed)
            with self.subTest(raw_event_resealed=change_raw_event), self.assertRaises(ValueError):
                validator.validate_application_case(changed)

    def test_working_pdf_terminal_protocol_and_zero_output_are_strict(self):
        for fmt in ("epub", "azw3"):
            case = synthetic_working(fmt)
            validator.validate_working_case(case)
            for field, value in (("version", 2), ("version", True), ("mode", "2"), ("code", "split_complete"),
                                 ("written_count", False), ("written_count", 1), ("exit_code", True), ("execution", {})):
                changed = deepcopy(case)
                changed["terminal_result"][field] = value
                reseal_working(changed)
                with self.subTest(fmt=fmt, field=field, value=value), self.assertRaisesRegex(ValueError, "terminal result"):
                    validator.validate_working_case(changed)

    def test_working_pdf_plan_requires_same_captured_source_and_complete_manual_coverage(self):
        for fmt in ("epub", "azw3"):
            case = synthetic_working(fmt)
            validator.validate_working_case(case)
            for mutation in ("source", "partial", "original", "destination", "working_identity"):
                changed = deepcopy(case)
                plan = changed["events"][-1]["plan"]
                if mutation == "source":
                    plan["source_identity"]["sha256"] = "f" * 64
                elif mutation == "partial":
                    plan.update(entries=plan["entries"][:1], ranges=[[0, 1]], coverage={"complete": False, "covered_pages": 1, "section_count": 1})
                elif mutation == "original":
                    plan["original_ebook_identity"]["sha256"] = "f" * 64
                elif mutation == "destination":
                    plan["output_naming"]["resolved_base"] = "C:/synthetic/foreign"
                else:
                    changed["terminal_result"]["diagnostic"]["working_pdf"]["sha256"] = "f" * 64
                canonical(changed["events"][-1])
                reseal_working(changed)
                with self.subTest(fmt=fmt, mutation=mutation), self.assertRaises(ValueError):
                    validator.validate_working_case(changed)

    def test_raw_console_or_duplicate_outcome_cannot_be_substituted(self):
        case = synthetic_cancel()
        for name in ("stdout", "stderr", "console_log"):
            changed = deepcopy(case)
            changed[name] += "altered bytes"
            with self.subTest(stream=name), self.assertRaises(ValueError):
                validator.validate_application_case(changed)
        case["stdout"] *= 2
        case["stdout_sha256"] = sha256(case["stdout"].encode()).hexdigest()
        with self.assertRaises(ValueError):
            validator.validate_application_case(case)

    def test_cancel_receipt_cannot_claim_a_successful_publication_or_complete_human_test(self):
        case = synthetic_cancel()
        case["outcome"].update(status="success", written_count=5, final_directory="C:/synthetic/published")
        reseal(case)
        with self.assertRaises(ValueError):
            validator.validate_application_case(case)
        with self.assertRaises(ValueError):
            validator.validate_ux_report({"schema_version": 1, "task_id": "M3-T04", "result": "UX_REGRESSION_PASSED",
                                         "success": True, "exit_code": 0, "human_or_viewer_opening_tested": True}, [], "")


if __name__ == "__main__":
    unittest.main()
