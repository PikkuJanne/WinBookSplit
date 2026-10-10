"""Strict, read-only AC-073 receipt validation; no deletion authority."""

from hashlib import sha256
import json
from pathlib import Path
import re


IDS = {"prior-repeat-1", "prior-repeat-2", "overlap-success", "overlap-mid-write", "repeat-after"}
PARTIAL_BYTES = b"WBS authored incomplete second PDF for AC-073\n"


def need(condition, message):
    if not condition:
        raise ValueError(message)


def digest(value):
    need(isinstance(value, str) and re.fullmatch(r"[0-9a-f]{64}", value), "Missing exact SHA-256")


def identity(value):
    need(isinstance(value, dict) and set(value) == {"device", "inode"}
         and type(value["device"]) is int and value["device"] >= 0
         and type(value["inode"]) is int and value["inode"] > 0, "Missing actual object identity")


def file(value):
    need(isinstance(value, dict) and set(value) == {"device", "inode", "size_bytes", "sha256"}, "Unexpected file observation")
    identity({key: value[key] for key in ("device", "inode")})
    need(type(value["size_bytes"]) is int and value["size_bytes"] > 0, "Empty observed file")
    digest(value["sha256"])


def snapshot(value):
    need(isinstance(value, dict) and set(value) == {"path", "directory_identity", "members"}
         and isinstance(value["path"], str) and Path(value["path"]).is_absolute(), "Invalid flat snapshot")
    identity(value["directory_identity"])
    need(isinstance(value["members"], dict) and bool(value["members"]), "Empty observed flat directory")
    for name, item in value["members"].items():
        need(isinstance(name, str) and name == Path(name).name and name not in {".", ".."}, "Nonliteral flat member")
        file(item)


def text_member(observation, name, text):
    need(isinstance(text, str), "Missing exact marker/manifest text")
    data = text.encode("utf-8")
    member = observation["members"].get(name, {})
    need(member.get("sha256") == sha256(data).hexdigest() and member.get("size_bytes") == len(data), "Observed text bytes differ from file hash")
    return json.loads(text)


def process(row, expected_exit):
    need(type(row.get("pid")) is int and row["pid"] > 0 and type(row.get("exit_code")) is int
         and row["exit_code"] == expected_exit and isinstance(row.get("command"), list) and row["command"]
         and all(isinstance(token, str) for token in row["command"]), "Missing actual native child command/status")
    need(type(row.get("launcher_pid")) is int and row["launcher_pid"] > 0
         and row.get("job_assigned_before_resume") is True and row.get("worker_membership_verified") is True
         and row.get("worker_handle_retained") is True and row.get("owned_tree_stopped") is True
         and type(row.get("job_active_after_exit")) is int and row["job_active_after_exit"] == 0
         and row.get("both_streams_eof") is True
         and type(row.get("worker_exit_code")) is int and row["worker_exit_code"] == expected_exit,
         "Actual interpreter/launcher membership, native exit or tree stop missing")
    need(row.get("stdout_truncated") is False and row.get("stderr_truncated") is False
         and isinstance(row.get("stdout"), str) and row.get("stderr") == "", "Incomplete or unexpected process streams")
    for stream in ("stdout", "stderr"):
        data = row[stream].encode("utf-8")
        need(row.get(stream + "_sha256") == sha256(data).hexdigest()
             and row.get(stream + "_bytes") == len(data) and row.get(stream + "_read_error", "missing") is None,
             "Raw process stream hash/size/error mismatch")
    frames = [json.loads(line) for line in row["stdout"].splitlines() if line.startswith("{")]
    need(len(frames) == 1 and frames[0] == row.get("result"), "Exactly one actual terminal result required")
    result = frames[0]
    need(result.get("protocol") == "winbooksplit.result" and type(result.get("version")) is int
         and result["version"] == 1 and result.get("exit_code") == expected_exit
         and type(result["exit_code"]) is int and result.get("mode") == "manual", "Terminal protocol/mode/native status differs")
    need(isinstance(row.get("started_at"), str) and bool(row["started_at"])
         and isinstance(row.get("completed_at"), str) and bool(row["completed_at"]), "Actual process time labels missing")
    return result


def plan(value, fixture):
    need(isinstance(value, dict) and value.get("mode") == "manual" and value.get("total_pages") == 10
         and value.get("ranges") == [[0, 3], [3, 6], [6, 10]]
         and value.get("coverage") == {"complete": True, "covered_pages": 10, "section_count": 3}, "Unexpected AC-073 full-page plan")
    entries = value.get("entries")
    need(isinstance(entries, list) and len(entries) == 3
         and [(entry.get("sequence"), entry.get("start"), entry.get("end")) for entry in entries]
             == [(1, 0, 3), (2, 3, 6), (3, 6, 10)], "Plan entries do not partition the source")
    names = [entry.get("filename") for entry in entries]
    need(all(isinstance(name, str) and name == Path(name).name and name.endswith(".pdf") for name in names)
         and len({name.casefold() for name in names}) == 3, "Unsafe or colliding planned filenames")
    source = value.get("source_identity", {})
    observed = fixture["file"]
    need(source.get("path") == fixture["path"] and source.get("sha256") == observed["sha256"]
         and source.get("size_bytes") == observed["size_bytes"] and source.get("binding") == "reader_snapshot", "Plan not bound to actual immutable source")
    return entries


def success(row, fixture, base):
    result = process(row, 0)
    need(result.get("status") == "success" and result.get("code") == "split_complete"
         and type(result.get("written_count")) is int and result["written_count"] == 3, "Successful frame count/category differs")
    entries = plan(result.get("plan"), fixture)
    execution = result.get("execution")
    need(isinstance(execution, dict) and execution.get("status") == "complete"
         and execution.get("mode") == "manual" and execution.get("written_count") == 3
         and type(execution["written_count"]) is int and execution.get("total_pages") == 10
         and execution.get("coverage") == result["plan"]["coverage"]
         and execution.get("source_identity") == result["plan"]["source_identity"], "Writer not bound to the displayed source/plan")
    run_id = execution.get("run_id", "")
    need(isinstance(run_id, str) and re.fullmatch(r"[0-9a-f]{32}", run_id), "Invalid successful run identity")
    final = Path(execution.get("final_directory", ""))
    need(final.is_absolute() and final.parent == Path(base) and final.name.endswith("_" + run_id)
         and "_20261010-120000_" in final.name, "Successful publication escaped the exact fixed-time base")
    publication = row.get("publication", {})
    observed = publication.get("snapshot")
    snapshot(observed)
    need(observed["path"] == str(final), "Publication snapshot identifies another directory")
    owner = text_member(observed, ".WinBookSplit-owner.json", publication.get("owner_text"))
    need(owner == {"schema_version": 1, "kind": "run", "run_id": run_id}, "Published owner not authenticated")
    manifest = text_member(observed, "WinBookSplit_Manifest.json", publication.get("manifest_text"))
    expected = {"schema_version": 1, "status": "complete", **{key: execution.get(key) for key in
                ("run_id", "final_directory", "mode", "total_pages", "source_identity", "coverage", "written_count", "outputs")}}
    need(manifest == expected and execution.get("manifest") == manifest
         and execution.get("manifest_filename") == "WinBookSplit_Manifest.json", "Manifest/terminal execution differs")
    outputs, reopened = execution.get("outputs"), publication.get("reopened")
    need(isinstance(outputs, list) and isinstance(reopened, list) and len(outputs) == len(reopened) == 3, "Reopened output records missing")
    need(set(observed["members"]) == {".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", *(entry["filename"] for entry in entries)}, "Publication has unknown members")
    for entry, written, actual in zip(entries, outputs, reopened):
        need(all(written.get(key) == value for key, value in entry.items()), "Writer changed prepared entry")
        first, last, name = entry["start"], entry["end"], entry["filename"]
        need(actual.get("filename") == name and actual.get("page_ids") == fixture["page_ids"][first:last]
             and actual.get("page_content_sha256") == fixture["page_content_sha256"][first:last]
             and written.get("page_count") == last - first and type(written["page_count"]) is int
             and written.get("sha256") == observed["members"][name]["sha256"]
             and written.get("size_bytes") == observed["members"][name]["size_bytes"], "Actual slice bytes/pages differ from immutable source")
    return execution


def readiness(value, row, fixture, base, nonce):
    need(isinstance(value, dict) and value.get("protocol") == "winbooksplit.test.isolation"
         and type(value.get("version")) is int and value["version"] == 1 and value.get("nonce") == nonce
         and value.get("role") == row["id"] and value.get("pid") == row["pid"]
         and value.get("phase") == "first-slice-held" and value.get("completed_real_slices") == 1
         and type(value["completed_real_slices"]) is int, "No actual held first slice from this child")
    entries = plan(row["result"]["plan"], fixture)
    stage = Path(value.get("stage", ""))
    run_id = value.get("run_id", "")
    need(re.fullmatch(r"[0-9a-f]{32}", run_id) and stage.parent == Path(base)
         and stage.name == ".WinBookSplit-stage-" + run_id, "Held stage ownership escaped the exact base")
    observed = row.get("held_snapshot")
    snapshot(observed)
    need(observed["path"] == str(stage) and observed["directory_identity"] == value.get("stage_identity"), "Held stage identity differs")
    names = {".WinBookSplit-owner.json", entries[0]["filename"]}
    need(set(observed["members"]) == set(value.get("ledger", {})) == names, "Held stage has unregistered or missing members")
    for name in names:
        need(value["ledger"][name] == {key: observed["members"][name][key] for key in ("device", "inode")}, "Held stage file identity differs from OutputRun ledger")
    owner = text_member(observed, ".WinBookSplit-owner.json", row.get("held_owner_text"))
    need(owner == {"schema_version": 1, "kind": "run", "run_id": run_id}, "Held owner marker differs")
    first = value.get("first_slice", {})
    observed_first = observed["members"][entries[0]["filename"]]
    need(first.get("filename") == entries[0]["filename"] and first.get("start") == 0 and first.get("end") == 3
         and first.get("page_count") == 3 and first.get("sha256") == observed_first["sha256"]
         and first.get("size_bytes") == observed_first["size_bytes"]
         and row.get("held_page_ids") == fixture["page_ids"][:3]
         and row.get("held_page_content_sha256") == fixture["page_content_sha256"][:3], "Barrier precedes a complete reopened first PDF")
    return entries


def validate_isolation_report(report, expected_source=None, *, cleanup_complete=True):
    need(isinstance(report, dict) and report.get("schema_version") == 1 and type(report["schema_version"]) is int
         and report.get("task_id") == "M4-T03" and report.get("result") == "RUN_ISOLATION_REGRESSION_PASSED"
         and report.get("success") is True and type(report.get("exit_code")) is int and report["exit_code"] == 0
         and report.get("acceptance_ids") == ["AC-073"] and report.get("source_unchanged") is True, "Not a passing scoped isolation receipt")
    source = report.get("tested_path_sha256")
    need(isinstance(source, dict) and bool(source) and source == report.get("source_before") == report.get("source_after"), "Tested source bytes changed")
    if expected_source is not None:
        need(source == expected_source, "Isolation source differs from enclosing gate")
    for value in source.values():
        digest(value)
    fixture = report.get("fixture", {})
    need(isinstance(fixture.get("path"), str) and Path(fixture["path"]).is_absolute()
         and fixture.get("page_ids") == list(range(1, 11)) and len(fixture.get("page_content_sha256", [])) == 10,
         "Missing ordinary authored ten-page family")
    file(fixture.get("file"))
    for value in fixture["page_content_sha256"]:
        digest(value)
    digest(fixture.get("provenance_sha256"))
    provenance_text = fixture.get("provenance_text")
    need(isinstance(provenance_text, str) and sha256(provenance_text.encode("utf-8")).hexdigest() == fixture["provenance_sha256"], "Authored provenance raw bytes differ")
    provenance = json.loads(provenance_text)
    need(provenance == fixture.get("provenance") and provenance.get("schema_version") == 1
         and set(provenance.get("fixtures", {})) == {"simple10"}, "Unexpected authored fixture family")
    authored = provenance["fixtures"]["simple10"]
    need(authored.get("filename") == Path(fixture["path"]).name and authored.get("sha256") == fixture["file"]["sha256"]
         and authored.get("size_bytes") == fixture["file"]["size_bytes"] and authored.get("page_ids") == fixture["page_ids"], "Provenance does not bind actual immutable input")
    protected = report.get("protected_before")
    need(isinstance(protected, dict) and set(protected) == {"source", "source_neighbor", "output_neighbor"}, "Incomplete protected file scope")
    for value in protected.values():
        need(isinstance(value, dict) and isinstance(value.get("path"), str) and Path(value["path"]).is_absolute(), "Protected literal path missing")
        file(value.get("file"))
    need(protected["source"] == {"path": fixture["path"], "file": fixture["file"]}
         and report.get("protected_after_failure") == protected and report.get("protected_after_all") == protected, "Input/neighbor identity or bytes changed")
    rows = report.get("cases")
    need(isinstance(rows, list) and len(rows) == 5 and {row.get("id") for row in rows} == IDS, "All five actual repeat/overlap roles required")
    by_id = {row["id"]: row for row in rows}
    base = report.get("output_base")
    need(isinstance(base, str) and Path(base).is_absolute(), "Exact shared output base missing")
    successes = [success(by_id[name], fixture, base) for name in ("prior-repeat-1", "prior-repeat-2", "overlap-success", "repeat-after")]
    need(len({row["run_id"] for row in successes}) == len({row["final_directory"] for row in successes}) == 4, "Repeated/concurrent publication reused ownership")
    priors = {name: by_id[name]["publication"]["snapshot"] for name in ("prior-repeat-1", "prior-repeat-2")}
    need(report.get("prior_after_second") == {"prior-repeat-1": priors["prior-repeat-1"]}
         and report.get("prior_before_overlap") == priors and report.get("prior_after_failure") == priors
         and report.get("prior_after_all") == priors, "Prior run hashes/identities changed")
    overlap = report.get("overlap", {})
    nonce = overlap.get("nonce", "")
    need(isinstance(nonce, str) and re.fullmatch(r"[0-9a-f]{32}", nonce), "Missing authored overlap control nonce")
    good, bad = by_id["overlap-success"], by_id["overlap-mid-write"]
    environment = report.get("environment", {})
    need(environment.get("python_version") == "3.14.8" and environment.get("pypdf_version") == "6.19.0"
         and isinstance(environment.get("python_executable"), str) and Path(environment["python_executable"]).is_absolute(), "Pinned actual interpreter/parser identity missing")
    work = Path(report.get("workspace", ""))
    engine, worker = report.get("engine_path"), report.get("worker_path")
    need(work.is_absolute() and Path(base) == work / "output"
         and isinstance(engine, str) and Path(engine).is_absolute() and isinstance(worker, str) and Path(worker).is_absolute(), "Exact owned worker/base paths missing")
    for row in rows:
        control = str(work / "controls" / row["id"])
        expected = [environment["python_executable"], "-I", "-B", "-X", "utf8", worker, "--engine", engine,
                    "--source", fixture["path"], "--base", base, "--control", control, "--role", row["id"], "--nonce", nonce]
        need(row.get("command") == expected and row.get("control_directory") == control and row.get("cwd") == str(work), "Native command not bound to exact fixed worker/input/role")
        boot = row.get("interpreter_boot", {})
        need(boot.get("protocol") == "winbooksplit.test.isolation" and type(boot.get("version")) is int and boot["version"] == 1
             and boot.get("nonce") == nonce and boot.get("role") == row["id"] and boot.get("phase") == "interpreter-held"
             and boot.get("pid") == row["pid"] and type(boot.get("parent_pid")) is int and boot["parent_pid"] > 0
             and (row["pid"] == row["launcher_pid"] or boot["parent_pid"] == row["launcher_pid"])
             and all(boot.get(key) == environment.get(key) for key in
                     ("python_executable", "python_version", "python_prefix", "python_base_prefix", "pypdf_path", "pypdf_version")),
             "Actual boot interpreter/package/job identity differs")
        need(row["result"]["plan"] == good["result"]["plan"], "Repeated/overlap prepared plans differ")
    completed = {name: by_id[name]["publication"]["snapshot"] for name in ("prior-repeat-1", "prior-repeat-2", "overlap-success")}
    need(report.get("completed_before_last_repeat") == completed and report.get("completed_after_last_repeat") == completed, "Later repeat changed the concurrent completed run")
    need(good["pid"] != bad["pid"] and overlap.get("live_pids_before_release") == [good["pid"], bad["pid"]]
         and overlap.get("success_live_after_failure") == good["pid"], "Actual simultaneous live children not proved")
    timeline = [overlap.get(key) for key in ("both_held_ns", "failure_release_ns", "failure_exit_ns", "held_check_ns", "success_release_ns", "success_exit_ns")]
    need(all(type(value) is int and value > 0 for value in timeline)
         and all(left < right for left, right in zip(timeline, timeline[1:])), "Overlap/release/cleanup timeline is impossible")
    bad_result = process(bad, 6)
    need(bad_result.get("status") == "write_error" and bad_result.get("code") == "output_write_failed"
         and bad_result.get("written_count") == 0 and type(bad_result["written_count"]) is int
         and bad_result.get("execution", "missing") is None and bad_result.get("fallback_modes") == [], "Mid-write fault falsely succeeded or changed category")
    entries = readiness(bad.get("ready"), bad, fixture, base, nonce)
    readiness(good.get("ready"), good, fixture, base, nonce)
    need(good["ready"]["run_id"] == successes[2]["run_id"]
         and bad["ready"]["run_id"] not in {item["run_id"] for item in successes}, "Stage and publication identities disagree")
    need(overlap.get("success_held_after_failure") == good["held_snapshot"]
         and overlap.get("failure_stage_absent") is True, "Failure cleanup changed another stage or retained its own")
    partial = bad.get("partial", {})
    second = entries[1]["filename"]
    need(partial.get("protocol") == "winbooksplit.test.isolation" and partial.get("version") == 1
         and partial.get("nonce") == nonce and partial.get("role") == bad["id"] and partial.get("pid") == bad["pid"]
         and partial.get("phase") == "partial-second-write" and partial.get("run_id") == bad["ready"]["run_id"]
         and partial.get("stage") == bad["ready"]["stage"] and partial.get("stage_identity") == bad["ready"]["stage_identity"]
         and partial.get("completed_real_slices") == 1 and partial.get("write_call") == 2
         and partial.get("path") == str(Path(partial["stage"]) / second)
         and partial.get("bytes_written") == partial.get("size_bytes") == len(PARTIAL_BYTES)
         and partial.get("sha256") == sha256(PARTIAL_BYTES).hexdigest(), "Fault not proven inside a registered real second stream after the first PDF")
    ledger = partial.get("ledger", {})
    need(set(ledger) == set(bad["ready"]["ledger"]) | {second}
         and all(ledger[name] == item for name, item in bad["ready"]["ledger"].items())
         and ledger[second] == {key: partial.get(key) for key in ("device", "inode")}, "Partial PDF not bound to the actual creation ledger")
    identity(ledger[second])
    diagnostic = bad_result.get("diagnostic", {})
    failed = bad.get("failure", {})
    observed = failed.get("snapshot")
    snapshot(observed)
    need(diagnostic.get("run_id") == bad["ready"]["run_id"] and diagnostic.get("cleanup_complete") is True
         and diagnostic.get("retained_staging", "missing") is None and diagnostic.get("cleanup_error", "missing") is None
         and diagnostic.get("record_path") == str(Path(observed["path"]) / "failure.json")
         and Path(observed["path"]).parent == Path(base)
         and Path(observed["path"]).name == ".WinBookSplit-failed-" + diagnostic["run_id"]
         and set(observed["members"]) == {".WinBookSplit-owner.json", "failure.json"}, "Failure ownership/cleanup evidence missing")
    owner = text_member(observed, ".WinBookSplit-owner.json", failed.get("owner_text"))
    record = text_member(observed, "failure.json", failed.get("record_text"))
    need(owner == {"schema_version": 1, "kind": "failed", "run_id": diagnostic["run_id"]}
         and record.get("schema_version") == 1 and record.get("status") == "failed"
         and record.get("code") == bad_result["code"] and record.get("run_id") == diagnostic["run_id"]
         and record.get("cleanup_complete") is True and record.get("retained_staging", "missing") is None
         and record.get("cleanup_error", "missing") is None, "Separately marked bounded failure record contradicts cleanup")
    expected_directories = {Path(row["final_directory"]).name for row in successes} | {Path(observed["path"]).name}
    need(report.get("base_directories_before_cleanup") == sorted(expected_directories), "Failure published output or other staging was removed")
    cleanup = report.get("cleanup", {})
    expected_snapshots = [by_id[name]["publication"]["snapshot"] for name in ("prior-repeat-1", "prior-repeat-2", "overlap-success", "repeat-after")] + [observed]
    need(cleanup.get("authenticated_targets") == expected_snapshots, "Owned cleanup must authenticate exactly these five directories")
    if cleanup_complete:
        need(cleanup.get("removed_paths") == [item["path"] for item in expected_snapshots]
             and cleanup.get("remaining_members") == ["synthetic-neighbor.txt"] and report.get("owned_output_cleanup") is True,
             "Actual owned cleanup/removal missing")
        need(report.get("protected_after_cleanup") == protected, "Owned cleanup changed an input or neighbor")
    else:
        need(cleanup.get("removed_paths") == [] and report.get("owned_output_cleanup") is False,
             "Pre-cleanup validation must not claim deletion")
