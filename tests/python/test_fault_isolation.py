"""Independently resealed contradictions in the AC-073 isolation receipt.

The positive model is synthetic receipt data, not native/PDF evidence. Actual
five-process execution is the separate characterize_isolation acceptance route.
"""

from copy import deepcopy
from hashlib import sha256
import importlib.util
import json
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import Mock


ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("wbs_isolation_receipt_test", ROOT / "tests/faults/validate_isolation.py")
validator = importlib.util.module_from_spec(spec)
spec.loader.exec_module(validator)


def encoded(value):
    return json.dumps(value, sort_keys=True, ensure_ascii=True)


def observed(data, inode):
    return {"device": 1, "inode": inode, "size_bytes": len(data), "sha256": sha256(data).hexdigest()}


def snapshot(path, texts, inode):
    return {"path": str(path), "directory_identity": {"device": 1, "inode": inode},
            "members": {name: observed(text.encode("utf-8"), inode * 10 + index)
                        for index, (name, text) in enumerate(texts.items(), 1)}}


def reseal(row):
    row["stdout"] = "Authored stream label\n" + encoded(row["result"]) + "\n"
    row["stderr"] = ""
    for name in ("stdout", "stderr"):
        data = row[name].encode("utf-8")
        row[name + "_sha256"], row[name + "_bytes"] = sha256(data).hexdigest(), len(data)
        row[name + "_truncated"] = False
        row[name + "_read_error"] = None


def model():
    # Paths are literal authored labels. No file is created or read here.
    drive = Path(ROOT.anchor)
    work = drive / "synthetic-isolation"
    base, source = work / "output", work / "fixtures/simple-10.pdf"
    fixture = {"path": str(source), "file": observed(b"synthetic-source-label", 10),
               "page_ids": list(range(1, 11)), "page_content_sha256": [sha256(str(i).encode()).hexdigest() for i in range(10)],
               "provenance_sha256": sha256(b"synthetic-provenance-label").hexdigest()}
    fixture["provenance"] = {"schema_version": 1, "fixtures": {"simple10": {"filename": source.name,
        "sha256": fixture["file"]["sha256"], "size_bytes": fixture["file"]["size_bytes"], "page_ids": fixture["page_ids"]}}}
    fixture["provenance_text"] = encoded(fixture["provenance"])
    fixture["provenance_sha256"] = sha256(fixture["provenance_text"].encode()).hexdigest()
    identity = {"path": str(source), "binding": "reader_snapshot", "sha256": fixture["file"]["sha256"],
                "size_bytes": fixture["file"]["size_bytes"]}
    entries = [{"sequence": i, "title": f"Section {i}", "start": start, "end": end,
                "parent_id": None, "reason": "manual", "filename": f"{i:02d} - Section {i}.pdf", "warnings": []}
               for i, (start, end) in enumerate(((0, 3), (3, 6), (6, 10)), 1)]
    plan = {"mode": "manual", "total_pages": 10, "ranges": [[0, 3], [3, 6], [6, 10]], "entries": entries,
            "source_identity": identity, "coverage": {"complete": True, "covered_pages": 10, "section_count": 3}}
    nonce = "a" * 32
    rows = []
    for index, role in enumerate(("prior-repeat-1", "prior-repeat-2", "overlap-success", "overlap-mid-write", "repeat-after"), 1):
        run_id = str(index) * 32
        row = {"id": role, "pid": 100 + index, "exit_code": 6 if role == "overlap-mid-write" else 0,
               "command": ["synthetic-pinned-python", role], "started_at": "2026-10-10T12:00:00Z", "completed_at": "2026-10-10T12:00:01Z"}
        result = {"protocol": "winbooksplit.result", "version": 1, "mode": "manual", "plan": deepcopy(plan),
                  "fallback_modes": [], "diagnostic": None, "warnings": []}
        if role == "overlap-mid-write":
            directory = base / (".WinBookSplit-failed-" + run_id)
            diagnostic = {"run_id": run_id, "cleanup_complete": True, "retained_staging": None,
                          "cleanup_error": None, "record_path": str(directory / "failure.json")}
            record = {"schema_version": 1, "status": "failed", "code": "output_write_failed", **diagnostic}
            owner = {"schema_version": 1, "kind": "failed", "run_id": run_id}
            text = {".WinBookSplit-owner.json": encoded(owner), "failure.json": encoded(record)}
            row["failure"] = {"snapshot": snapshot(directory, text, 500), "owner_text": text[".WinBookSplit-owner.json"], "record_text": text["failure.json"]}
            result.update(status="write_error", code="output_write_failed", exit_code=6, written_count=0, execution=None, diagnostic=diagnostic)
        else:
            final = base / ("Book_20261010-120000_" + run_id)
            outputs = [{**entry, "page_count": entry["end"] - entry["start"],
                        "sha256": sha256((role + entry["filename"]).encode()).hexdigest(),
                        "size_bytes": len((role + entry["filename"]).encode())} for entry in entries]
            manifest = {"schema_version": 1, "status": "complete", "run_id": run_id, "final_directory": str(final),
                        "mode": "manual", "total_pages": 10, "source_identity": identity, "coverage": plan["coverage"],
                        "written_count": 3, "outputs": outputs}
            execution = {**manifest, "manifest_filename": "WinBookSplit_Manifest.json", "manifest": manifest}
            owner = {"schema_version": 1, "kind": "run", "run_id": run_id}
            text = {".WinBookSplit-owner.json": encoded(owner), "WinBookSplit_Manifest.json": encoded(manifest),
                    **{entry["filename"]: role + entry["filename"] for entry in entries}}
            row["publication"] = {"snapshot": snapshot(final, text, index * 100), "owner_text": text[".WinBookSplit-owner.json"],
                                  "manifest_text": text["WinBookSplit_Manifest.json"],
                                  "reopened": [{"filename": entry["filename"], "page_ids": fixture["page_ids"][entry["start"]:entry["end"]],
                                                "page_content_sha256": fixture["page_content_sha256"][entry["start"]:entry["end"]]} for entry in entries]}
            result.update(status="success", code="split_complete", exit_code=0, written_count=3, execution=execution)
        row["result"] = result
        if role in {"overlap-success", "overlap-mid-write"}:
            stage = base / (".WinBookSplit-stage-" + run_id)
            first_name = entries[0]["filename"]
            text = {".WinBookSplit-owner.json": encoded({"schema_version": 1, "kind": "run", "run_id": run_id}), first_name: role + first_name}
            held = snapshot(stage, text, index * 100 + 1)
            ledger = {name: {key: value[key] for key in ("device", "inode")} for name, value in held["members"].items()}
            row.update(held_snapshot=held, held_owner_text=text[".WinBookSplit-owner.json"], held_page_ids=fixture["page_ids"][:3],
                       held_page_content_sha256=fixture["page_content_sha256"][:3],
                       ready={"protocol": "winbooksplit.test.isolation", "version": 1, "nonce": nonce, "role": role,
                              "pid": row["pid"], "phase": "first-slice-held", "run_id": run_id, "stage": str(stage),
                              "stage_identity": held["directory_identity"], "completed_real_slices": 1, "ledger": ledger,
                              "first_slice": {"filename": first_name, "start": 0, "end": 3, "page_count": 3,
                                              "sha256": held["members"][first_name]["sha256"], "size_bytes": held["members"][first_name]["size_bytes"]}})
            if role == "overlap-mid-write":
                second = entries[1]["filename"]
                row["partial"] = {**{key: row["ready"][key] for key in ("protocol", "version", "nonce", "role", "pid", "run_id", "stage", "stage_identity")},
                                  "phase": "partial-second-write", "completed_real_slices": 1, "write_call": 2,
                                  "path": str(stage / second), "device": 1, "inode": 4444,
                                  "size_bytes": len(validator.PARTIAL_BYTES), "bytes_written": len(validator.PARTIAL_BYTES),
                                  "sha256": sha256(validator.PARTIAL_BYTES).hexdigest(),
                                  "ledger": {**ledger, second: {"device": 1, "inode": 4444}}}
        reseal(row)
        rows.append(row)
    good, bad = rows[2:4]
    prior = {row["id"]: deepcopy(row["publication"]["snapshot"]) for row in rows[:2]}
    protected = {"source": {"path": str(source), "file": fixture["file"]},
                 "source_neighbor": {"path": str(work / "synthetic-input-neighbor.txt"), "file": observed(b"input neighbor", 20)},
                 "output_neighbor": {"path": str(base / "synthetic-neighbor.txt"), "file": observed(b"output neighbor", 30)}}
    targets = [rows[index]["publication"]["snapshot"] for index in (0, 1, 2, 4)] + [bad["failure"]["snapshot"]]
    source_map = {"engine/winbooksplit_engine.py": sha256(b"synthetic-source-map").hexdigest()}
    engine, worker, python = str(ROOT / "engine/winbooksplit_engine.py"), str(ROOT / "tests/faults/worker.py"), str(drive / "pinned-python.exe")
    for row in rows:
        control = str(work / "controls" / row["id"])
        row.update(control_directory=control, cwd=str(work), command=[python, "-I", "-B", "-X", "utf8", worker,
            "--engine", engine, "--source", str(source), "--base", str(base), "--control", control, "--role", row["id"], "--nonce", nonce])
        row.update(launcher_pid=row["pid"] + 1000, job_assigned_before_resume=True, worker_membership_verified=True,
                   worker_handle_retained=True, owned_tree_stopped=True, worker_exit_code=row["exit_code"], job_active_after_exit=0, both_streams_eof=True,
                   interpreter_boot={"protocol": "winbooksplit.test.isolation", "version": 1, "phase": "interpreter-held",
                       "nonce": nonce, "role": row["id"], "pid": row["pid"], "parent_pid": row["pid"] + 1000,
                       "python_executable": python, "python_version": "3.14.8", "python_prefix": str(drive / "synthetic-venv"),
                       "python_base_prefix": str(drive / "synthetic-base"), "pypdf_version": "6.19.0", "pypdf_path": str(drive / "synthetic-pypdf/__init__.py")})
    return {"schema_version": 1, "task_id": "M4-T03", "result": "RUN_ISOLATION_REGRESSION_PASSED", "success": True,
            "exit_code": 0, "acceptance_ids": ["AC-073"], "source_unchanged": True, "tested_path_sha256": source_map,
            "source_before": deepcopy(source_map), "source_after": deepcopy(source_map), "fixture": fixture,
            "workspace": str(work), "engine_path": engine, "worker_path": worker,
            "environment": {key: rows[0]["interpreter_boot"][key] for key in
                ("python_executable", "python_version", "python_prefix", "python_base_prefix", "pypdf_path", "pypdf_version")},
            "output_base": str(base), "cases": rows, "protected_before": protected, "protected_after_failure": deepcopy(protected),
            "protected_after_all": deepcopy(protected), "protected_after_cleanup": deepcopy(protected),
            "prior_after_second": {"prior-repeat-1": deepcopy(prior["prior-repeat-1"])}, "prior_before_overlap": prior,
            "prior_after_failure": deepcopy(prior), "prior_after_all": deepcopy(prior),
            "completed_before_last_repeat": {**deepcopy(prior), "overlap-success": deepcopy(good["publication"]["snapshot"])},
            "completed_after_last_repeat": {**deepcopy(prior), "overlap-success": deepcopy(good["publication"]["snapshot"])},
            "overlap": {"nonce": nonce, "live_pids_before_release": [good["pid"], bad["pid"]], "success_live_after_failure": good["pid"],
                        "both_held_ns": 10, "failure_release_ns": 20, "failure_exit_ns": 30, "held_check_ns": 40,
                        "success_release_ns": 50, "success_exit_ns": 60, "failure_stage_absent": True,
                        "success_held_after_failure": deepcopy(good["held_snapshot"])},
            "base_directories_before_cleanup": sorted(Path(item["path"]).name for item in targets),
            "owned_output_cleanup": True, "cleanup": {"authenticated_targets": targets,
                "removed_paths": [item["path"] for item in targets], "remaining_members": ["synthetic-neighbor.txt"]}}


class FaultIsolationReceiptTests(unittest.TestCase):
    def setUp(self):
        self.receipt = model()
        validator.validate_isolation_report(self.receipt)

    def reject(self, mutate):
        value = deepcopy(self.receipt)
        mutate(value)
        with self.assertRaises((ValueError, KeyError, TypeError)):
            validator.validate_isolation_report(value)

    def test_source_neighbor_and_prior_hash_or_identity_change_is_not_hidden_by_preservation_flags(self):
        for field in ("protected_after_failure", "protected_after_all", "protected_after_cleanup"):
            for name in ("source", "source_neighbor", "output_neighbor"):
                for key, replacement in (("sha256", "f" * 64), ("inode", 999)):
                    with self.subTest(field=field, name=name, key=key):
                        self.reject(lambda value: value[field][name]["file"].update({key: replacement}))
        self.reject(lambda value: value["prior_after_failure"]["prior-repeat-1"]["members"]["01 - Section 1.pdf"].update(sha256="e" * 64))
        self.reject(lambda value: value["source_after"].update({"engine/winbooksplit_engine.py": "d" * 64}))
        with self.assertRaises(ValueError):
            validator.validate_isolation_report(self.receipt, {"another-source.py": "c" * 64})

    def test_overlap_requires_live_distinct_children_real_first_slice_and_ordered_release(self):
        mutations = [lambda value: value["overlap"].update(live_pids_before_release=[103, 103]),
                     lambda value: value["overlap"].update(success_live_after_failure=None),
                     lambda value: value["overlap"].update(success_release_ns=25),
                     lambda value: value["cases"][2]["ready"].update(completed_real_slices=0),
                     lambda value: value["cases"][2].update(held_page_ids=[2, 3, 4]),
                     lambda value: value["cases"][3]["ready"].update(pid=999)]
        for index, mutate in enumerate(mutations):
            with self.subTest(index=index):
                self.reject(mutate)

    def test_venv_launcher_is_not_silently_treated_as_the_actual_worker(self):
        for flag in ("job_assigned_before_resume", "worker_membership_verified", "worker_handle_retained", "owned_tree_stopped", "both_streams_eof"):
            with self.subTest(flag=flag):
                self.reject(lambda value: value["cases"][2].update({flag: False}))
        self.reject(lambda value: value["cases"][2].update(worker_exit_code=1))
        self.reject(lambda value: value["cases"][2].update(job_active_after_exit=1))
        self.reject(lambda value: value["cases"][2]["interpreter_boot"].update(parent_pid=999))
        self.reject(lambda value: value["cases"][2]["interpreter_boot"].update(pypdf_path=str(ROOT / "shadow.py")))
        self.reject(lambda value: value["cases"][2].update(stdout_read_error="incomplete pipe"))

    def test_real_partial_second_stream_must_have_registered_identity_and_completed_first_slice(self):
        for field, replacement in (("write_call", 1), ("completed_real_slices", 0), ("size_bytes", 0),
                                   ("bytes_written", 0), ("inode", 999), ("nonce", "b" * 32)):
            with self.subTest(field=field):
                self.reject(lambda value: value["cases"][3]["partial"].update({field: replacement}))
        self.reject(lambda value: value["cases"][3]["partial"]["ledger"].pop("02 - Section 2.pdf"))

    def test_failure_cannot_publish_claim_success_or_remove_the_other_owned_stage(self):
        def forged_success(value):
            row = value["cases"][3]
            row["result"].update(status="success", code="split_complete", written_count=1)
            reseal(row)
        self.reject(forged_success)
        self.reject(lambda value: value["overlap"]["success_held_after_failure"]["directory_identity"].update(inode=999))
        self.reject(lambda value: value["overlap"]["success_held_after_failure"]["members"]["01 - Section 1.pdf"].update(sha256="a" * 64))
        self.reject(lambda value: value["overlap"].update(failure_stage_absent=False))
        self.reject(lambda value: value["base_directories_before_cleanup"].append("foreign-final"))

    def test_resealed_output_count_page_order_and_manifest_must_match_actual_complete_publication(self):
        def count(value):
            row = value["cases"][2]
            row["result"]["written_count"] = 1
            reseal(row)
        self.reject(count)
        self.reject(lambda value: value["cases"][2]["publication"]["reopened"][0].update(page_ids=[3, 2, 1]))
        self.reject(lambda value: value["cases"][2]["publication"]["reopened"][0].update(page_content_sha256=["b" * 64] * 3))
        self.reject(lambda value: value["cases"][2]["publication"]["snapshot"]["members"]["01 - Section 1.pdf"].update(sha256="b" * 64))

    def test_cleanup_requires_exact_previously_authenticated_targets_and_actual_removal(self):
        self.reject(lambda value: value["cleanup"]["authenticated_targets"][0].update(path=str(ROOT.anchor + "foreign")))
        self.reject(lambda value: value["cleanup"]["removed_paths"].pop())
        self.reject(lambda value: value["cleanup"].update(remaining_members=["synthetic-neighbor.txt", "foreign.txt"]))
        self.reject(lambda value: value.update(owned_output_cleanup=False))
        pending = deepcopy(self.receipt)
        pending["owned_output_cleanup"] = False
        pending["cleanup"]["removed_paths"] = []
        validator.validate_isolation_report(pending, cleanup_complete=False)
        with self.assertRaises(ValueError):
            validator.validate_isolation_report(pending)

    def test_first_shutdown_failure_does_not_skip_other_owned_children_or_replace_primary_error(self):
        spec = importlib.util.spec_from_file_location("wbs_isolation_recovery_test", ROOT / "tests/faults/characterize_isolation.py")
        helper = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(helper)
        first = SimpleNamespace(role="overlap-mid-write", stop=Mock(side_effect=OSError("authored first job close failure")),
                                observation=Mock(return_value={"owned_tree_stopped": False}))
        second = SimpleNamespace(role="overlap-success", stop=Mock(), observation=Mock())
        primary = ValueError("authored primary assertion")
        helper.recover_children([first, second], primary)
        first.stop.assert_called_once_with()
        second.stop.assert_called_once_with()
        self.assertEqual(str(primary), "authored primary assertion")
        self.assertEqual(len(primary.__notes__), 1)
        self.assertEqual(helper.LAST_CASE["owned_child_cleanup_errors"][0]["actual_child"], {"owned_tree_stopped": False})
        second.stop.reset_mock()
        with self.assertRaisesRegex(OSError, "retain workspace"):
            helper.recover_children([first, second], None)
        second.stop.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
