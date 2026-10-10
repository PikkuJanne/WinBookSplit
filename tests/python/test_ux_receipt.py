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
    case["stdout"] = "[OUTCOME] " + compact(case["outcome"]) + "\r\n"
    for name in ("stdout", "stderr"):
        case[name + "_sha256"] = sha256(case[name].encode()).hexdigest()


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
