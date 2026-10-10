"""AC-073: repeated real publications overlap a real mid-second-write failure.

Only authored ordinary PDFs are processed. The child interposer changes a fixed
timestamp and injects a stream-write error; planning, reservation, slice writing,
validation, publication and cleanup use the shipped engine unchanged.
"""

import argparse
import ctypes
from ctypes import wintypes
from datetime import datetime, timezone
from hashlib import sha256
import importlib.util
import json
import os
from pathlib import Path
import platform
import subprocess
import sys
import tempfile
import threading
import time
import uuid


ROOT = Path(__file__).resolve().parents[2]
ENGINE = ROOT / "engine/winbooksplit_engine.py"
WORKER = Path(__file__).with_name("worker.py")


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


output = load("wbs_fault_output_helpers", ROOT / "tests/output/characterize_output.py")
manual, history, plans, runner, require = output.manual, output.history, output.plans, output.runner, output.require
launchers = load("wbs_fault_owned_cleanup", ROOT / "tests/manual/current_launchers.py")
validator = load("wbs_fault_isolation_validator", Path(__file__).with_name("validate_isolation.py"))
jobs = load("wbs_fault_owned_job", ROOT / "engine/winbooksplit_job.py")
COMPLETED_CASES, LAST_CASE = [], {}


def file_observation(path):
    details = os.lstat(path)
    require(path.is_file() and not history.is_reparse(path), "Authored file became nonordinary")
    data = path.read_bytes()
    after = os.lstat(path)
    require((after.st_dev, after.st_ino, after.st_size) == (details.st_dev, details.st_ino, details.st_size)
            and not history.is_reparse(path) and len(data) == details.st_size, "Observed file identity/size changed")
    return {"device": details.st_dev, "inode": details.st_ino,
            "size_bytes": len(data), "sha256": sha256(data).hexdigest()}


def flat_snapshot(directory):
    require(directory.is_absolute() and directory.resolve(strict=True) == directory
            and directory.is_dir() and not history.is_reparse(directory), "Flat snapshot refuses escaped/reparse directory")
    details = directory.lstat()
    observed = {"path": str(directory), "directory_identity": {"device": details.st_dev, "inode": details.st_ino},
                "members": {path.name: file_observation(path) for path in sorted(directory.iterdir())}}
    require(directory.lstat().st_ino == details.st_ino
            and set(observed["members"]) == {path.name for path in directory.iterdir()}, "Flat snapshot changed during capture")
    return observed


def protected_snapshot(source, source_neighbor, output_neighbor):
    return {name: {"path": str(path), "file": file_observation(path)} for name, path in
            (("source", source), ("source_neighbor", source_neighbor), ("output_neighbor", output_neighbor))}


class Child:
    """One owned job, distinguishing Windows venv launcher and interpreter."""
    def __init__(self, role, source, base, control, nonce, work):
        self.role, self.control = role, control
        control.mkdir()
        self.command = [sys.executable, "-I", "-B", "-X", "utf8", str(WORKER),
                        "--engine", str(ENGINE), "--source", str(source), "--base", str(base),
                        "--control", str(control), "--role", role, "--nonce", nonce]
        self.started_at = datetime.now(timezone.utc).isoformat()
        self.job = jobs.JobProcess(self.command, cwd=work, env=history.clean_environment(work))
        self.process = self.job.process
        self.worker_handle, self.worker_pid, self.boot = None, None, None
        self.worker_membership, self.tree_stopped = False, False
        self.actual_worker_exit = None
        self.active_after_exit = None
        self.job._api.GetExitCodeProcess.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
        self.job._api.GetExitCodeProcess.restype = wintypes.BOOL
        self.streams = {name: {"data": bytearray(), "hash": sha256(), "count": 0, "error": None}
                        for name in ("stdout", "stderr")}
        self.threads = []
        for name in self.streams:
            thread = threading.Thread(target=self.drain, args=(name,), daemon=True)
            thread.start()
            self.threads.append(thread)
        try:
            deadline = time.monotonic() + 20
            while not (control / "started.json").is_file():
                require(self.process.poll() is None and time.monotonic() < deadline, "Owned interpreter boot handshake missing")
                time.sleep(0.01)
            self.boot = json.loads((control / "started.json").read_text(encoding="utf-8"))
            actual_pid = self.boot.get("pid")
            require(type(actual_pid) is int and actual_pid > 0 and self.boot.get("role") == role
                    and self.boot.get("nonce") == nonce and self.boot.get("phase") == "interpreter-held", "Owned interpreter boot identity differs")
            require(actual_pid == self.process.pid or self.boot.get("parent_pid") == self.process.pid,
                    "Authored interpreter is not the launched process or its venv redirector child")
            self.worker_handle = self.job._api.OpenProcess(0x1000 | 0x00100000, False, actual_pid)
            require(bool(self.worker_handle), "Cannot retain the authored interpreter handle")
            belongs = wintypes.BOOL()
            require(self.job._api.GetProcessId(self.worker_handle) == actual_pid
                    and self.job._api.IsProcessInJob(self.worker_handle, self.job._job, ctypes.byref(belongs))
                    and belongs.value and self.worker_status() == 259,
                    "Actual interpreter identity/live owned-job membership not proved")
            self.worker_pid, self.worker_membership = actual_pid, True
            require(self.boot.get("python_executable") == sys.executable and self.boot.get("python_prefix") == sys.prefix
                    and self.boot.get("python_base_prefix") == sys.base_prefix
                    and self.boot.get("python_version") == platform.python_version()
                    and self.boot.get("pypdf_path") == history.pypdf.__file__
                    and self.boot.get("pypdf_version") == history.pypdf.__version__, "Actual interpreter/pypdf provenance differs")
            atomic_gate(control, "identity-ack", nonce)
        except BaseException:
            self.stop()
            LAST_CASE["boot_failure_child"] = self.observation()
            raise

    def worker_status(self):
        status = wintypes.DWORD()
        require(self.worker_handle and self.job._api.GetExitCodeProcess(self.worker_handle, ctypes.byref(status)),
                "Retained interpreter exit query failed")
        return status.value

    def drain(self, name):
        state = self.streams[name]
        try:
            with getattr(self.process, name) as stream:
                while data := stream.read(4096):
                    state["hash"].update(data)
                    state["count"] += len(data)
                    available = max(0, 8 * 1024 * 1024 - len(state["data"]))
                    state["data"].extend(data[:available])
        except BaseException as error:
            state["error"] = repr(error)

    def observation(self):
        record = {"id": self.role, "command": self.command, "cwd": str(self.control.parent.parent),
                  "control_directory": str(self.control), "pid": self.worker_pid,
                  "launcher_pid": self.process.pid, "interpreter_boot": self.boot,
                  "job_assigned_before_resume": self.job.assigned_before_resume,
                  "worker_membership_verified": self.worker_membership,
                  "worker_handle_retained": self.worker_pid is not None,
                  "worker_exit_code": self.actual_worker_exit, "owned_tree_stopped": self.tree_stopped,
                  "job_active_after_exit": self.active_after_exit,
                  "both_streams_eof": not any(thread.is_alive() for thread in self.threads),
                  "exit_code": self.process.poll(), "started_at": self.started_at}
        for name, state in self.streams.items():
            record.update({name: bytes(state["data"]).decode("utf-8", errors="backslashreplace"),
                           name + "_sha256": state["hash"].hexdigest(), name + "_bytes": state["count"],
                           name + "_truncated": state["count"] != len(state["data"]), name + "_read_error": state["error"]})
        return record

    def finish(self, timeout=40):
        self.process.wait(timeout=timeout)
        self.job.wait_tree(timeout=5)
        self.tree_stopped = True
        self.active_after_exit = self.job.active_process_count()
        require(self.active_after_exit == 0, "Owned job retained a live process after exit")
        self.actual_worker_exit = self.worker_status()
        require(self.actual_worker_exit == self.process.returncode, "Actual interpreter and venv launcher exit statuses differ")
        for thread in self.threads:
            thread.join(timeout=5)
        require(not any(thread.is_alive() for thread in self.threads), "Owned child pipe EOF not observed")
        require(all(state["error"] is None and state["count"] == len(state["data"]) for state in self.streams.values()),
                "Owned child stream incomplete")
        record = self.observation()
        record["completed_at"] = datetime.now(timezone.utc).isoformat()
        record["result"] = manual.result_record(record)
        return record

    def stop(self):
        try:
            if not self.tree_stopped:
                if self.job.active_process_count():
                    self.job.terminate_tree()
                self.job.wait_tree(timeout=10)
                self.tree_stopped = True
            self.active_after_exit = self.job.active_process_count()
            self.process.wait(timeout=10)
            if self.worker_handle is not None:
                self.actual_worker_exit = self.worker_status()
            for thread in self.threads:
                thread.join(timeout=5)
            require(not any(thread.is_alive() for thread in self.threads), "Owned recovery pipe EOF not observed")
        finally:
            if self.worker_handle is not None:
                self.job._api.CloseHandle(self.worker_handle)
                self.worker_handle = None
            self.job.close()


def publication(row, source, base, generator):
    checked = output.successful_case(row["id"], source, base, row, generator)
    final = Path(checked["writer_result"]["final_directory"])
    row["publication"] = {"snapshot": flat_snapshot(final),
                          "owner_text": (final / ".WinBookSplit-owner.json").read_text(encoding="utf-8"),
                          "manifest_text": (final / "WinBookSplit_Manifest.json").read_text(encoding="utf-8"),
                          "reopened": [{"filename": item["filename"], "page_ids": item["page_ids"],
                                        "page_content_sha256": plans.page_content(final / item["filename"])}
                                       for item in checked["outputs"]]}
    return row


def wait_ready(children):
    deadline = time.monotonic() + 25
    while not all((child.control / "ready.json").is_file() for child in children):
        require(all(child.process.poll() is None for child in children), "Owned overlap child exited before its real first-slice barrier")
        require(time.monotonic() < deadline, "Owned overlap first-slice readiness deadline elapsed")
        time.sleep(0.01)
    ready = []
    for child in children:
        value = json.loads((child.control / "ready.json").read_text(encoding="utf-8"))
        require(value["pid"] == child.worker_pid and value["role"] == child.role, "Ready event not from exact owned interpreter")
        stage = Path(value["stage"])
        row = {**child.observation(), "ready": value, "held_snapshot": flat_snapshot(stage),
               "held_owner_text": (stage / ".WinBookSplit-owner.json").read_text(encoding="utf-8"),
               "held_page_ids": output.history.load_generator().page_ids(stage / value["first_slice"]["filename"]),
               "held_page_content_sha256": plans.page_content(stage / value["first_slice"]["filename"])}
        ready.append(row)
    return ready


def remove_authenticated(observed, marker):
    directory = Path(observed["path"])
    require(flat_snapshot(directory) == observed, "Owned cleanup refuses changed identity or bytes")
    launchers.remove_known_directory(directory, directory.parent, set(observed["members"]), ".WinBookSplit-owner.json", marker,
        expected_sha256={name: item["sha256"] for name, item in observed["members"].items()},
        expected_file_identities={name: {key: item[key] for key in ("device", "inode")} for name, item in observed["members"].items()},
        expected_directory_identity=observed["directory_identity"])


def atomic_gate(control, name, nonce):
    pending = control / (name + ".pending")
    with pending.open("xb") as stream:
        stream.write(nonce.encode("ascii"))
        stream.flush()
        os.fsync(stream.fileno())
    os.rename(pending, control / name)


def release(child, nonce):
    atomic_gate(child.control, "release", nonce)


def recover_children(children, primary_error):
    """Attempt every owned job even when an earlier shutdown/close fails."""
    errors = []
    for child in children:
        try:
            child.stop()
        except BaseException as error:
            try:
                observation = child.observation()
            except BaseException as observer_error:
                observation = {"id": child.role, "observation_error": repr(observer_error)}
            errors.append({"id": child.role, "error": repr(error), "actual_child": observation})
    if errors:
        LAST_CASE["owned_child_cleanup_errors"] = errors
        if primary_error is not None:
            primary_error.add_note("Owned child recovery unproved: " + repr(errors))
        else:
            raise OSError("Owned child shutdown/close unproved; retain workspace: " + repr(errors))


def characterize(work):
    global LAST_CASE
    COMPLETED_CASES.clear()
    LAST_CASE = {"phase": "initialization", "workspace": str(work)}
    require(os.name == "nt" and sys.flags.isolated and sys.dont_write_bytecode, "Actual isolated Windows developer Python required")
    require(work.is_absolute() and not work.resolve().is_relative_to(ROOT), "External authored workspace required")
    if not work.exists():
        require(work.parent.resolve(strict=True) == work.parent and not history.is_reparse(work.parent), "Ordinary owned workspace parent required")
        work.mkdir()
    require(work.is_absolute() and work.resolve(strict=True) == work and not history.is_reparse(work)
            and not any(work.iterdir()), "New empty ordinary owned isolation workspace required")
    before = runner.source_manifest()
    manual.trusted_original_sources()
    generator = history.load_generator()
    definitions = generator.FIXTURE_DEFINITIONS
    try:
        generator.FIXTURE_DEFINITIONS = {"simple10": definitions["simple10"]}
        fixtures = generator.generate_fixtures(work / "fixtures")
    finally:
        generator.FIXTURE_DEFINITIONS = definitions
    source = fixtures["simple10"]
    source_neighbor = work / "synthetic-input-neighbor.txt"
    source_neighbor.write_bytes(b"AC-073 authored input neighbor must remain unchanged\n")
    base, controls = work / "output", work / "controls"
    base.mkdir()
    controls.mkdir()
    output_neighbor = base / "synthetic-neighbor.txt"
    output_neighbor.write_bytes(b"AC-073 authored output neighbor must remain unchanged\n")
    protected = protected_snapshot(source, source_neighbor, output_neighbor)
    fixture = {"path": str(source), "file": protected["source"]["file"], "page_ids": generator.page_ids(source),
               "page_content_sha256": plans.page_content(source),
               "provenance": json.loads((source.parent / "provenance.json").read_text(encoding="utf-8")),
               "provenance_text": (source.parent / "provenance.json").read_text(encoding="utf-8"),
               "provenance_sha256": history.file_digest(source.parent / "provenance.json")}
    report = {"schema_version": 1, "task_id": "M4-T03", "acceptance_ids": ["AC-073"],
              "workspace": str(work), "output_base": str(base), "fixture": fixture,
              "engine_path": str(ENGINE), "worker_path": str(WORKER),
              "tested_path_sha256": before, "source_before": before, "protected_before": protected,
              "environment": {"python_executable": sys.executable, "python_version": platform.python_version(),
                              "python_prefix": sys.prefix, "python_base_prefix": sys.base_prefix, "pypdf_path": history.pypdf.__file__,
                              "pypdf_version": history.pypdf.__version__, "platform": platform.platform()},
              "limits": ["Authored ordinary PDF only; no private files, Explorer, human, Calibre or crash/power-loss claim.",
                         "Fixed test-only timestamp, first-slice gates and second-stream write fault; shipped source unchanged.",
                         "Readiness proves both owned processes hold real stages, not simultaneous CPU execution.",
                         "Ownership/cleanup observations do not sandbox arbitrary same-account filesystem mutation."]}
    git = history.run(["git", "rev-parse", "HEAD", "HEAD^{tree}"], ROOT, environment=history.clean_environment(work))
    require(git["exit_code"] == 0, "Cannot record actual source Git identity")
    report["git_commit"], report["git_tree"] = git["stdout"].splitlines()
    children, priors, nonce = [], {}, uuid.uuid4().hex
    primary_error = None
    try:
        for role in ("prior-repeat-1", "prior-repeat-2"):
            LAST_CASE = {"phase": role, "report_so_far": report}
            child = Child(role, source, base, controls / role, nonce, work)
            children.append(child)
            row = publication(child.finish(), source, base, generator)
            COMPLETED_CASES.append(row)
            if priors:
                report["prior_after_second"] = {name: flat_snapshot(Path(item["path"])) for name, item in priors.items()}
            priors[role] = row["publication"]["snapshot"]
            require(protected_snapshot(source, source_neighbor, output_neighbor) == protected, "Repeat modified input/neighbor")
        report["prior_before_overlap"] = {name: flat_snapshot(Path(item["path"])) for name, item in priors.items()}
        good = Child("overlap-success", source, base, controls / "overlap-success", nonce, work)
        children.append(good)
        bad = Child("overlap-mid-write", source, base, controls / "overlap-mid-write", nonce, work)
        children.append(bad)
        LAST_CASE = {"phase": "awaiting-real-overlap", "report_so_far": report}
        good_held, bad_held = wait_ready([good, bad])
        require(good.process.poll() is None and bad.process.poll() is None, "Real staged overlap not live at capture")
        overlap = {"nonce": nonce, "both_held_ns": time.monotonic_ns(),
                   "live_pids_before_release": [good.worker_pid, bad.worker_pid]}
        report["overlap"] = overlap
        LAST_CASE = {"phase": "both-real-stages-held", "held_rows": [good_held, bad_held], "report_so_far": report}
        overlap["failure_release_ns"] = time.monotonic_ns()
        release(bad, nonce)
        bad_row = {**bad_held, **bad.finish()}
        overlap["failure_exit_ns"] = time.monotonic_ns()
        require(bad_row["exit_code"] == 6, "Actual mid-write failure lost native6")
        bad_row["partial"] = json.loads((bad.control / "partial.json").read_text(encoding="utf-8"))
        diagnostic = bad_row["result"]["diagnostic"]
        failed_directory = Path(diagnostic["record_path"]).parent
        bad_row["failure"] = {"snapshot": flat_snapshot(failed_directory),
                              "owner_text": (failed_directory / ".WinBookSplit-owner.json").read_text(encoding="utf-8"),
                              "record_text": (failed_directory / "failure.json").read_text(encoding="utf-8")}
        COMPLETED_CASES.append(bad_row)
        require(good.process.poll() is None, "Other held run exited before failed-run cleanup was observed")
        require(good.worker_status() == 259, "Other actual interpreter no longer live after failed tree shutdown")
        overlap["success_live_after_failure"] = good.worker_pid
        overlap["failure_stage_absent"] = not os.path.lexists(bad_held["ready"]["stage"])
        overlap["success_held_after_failure"] = flat_snapshot(Path(good_held["ready"]["stage"]))
        overlap["held_check_ns"] = time.monotonic_ns()
        report["prior_after_failure"] = {name: flat_snapshot(Path(item["path"])) for name, item in priors.items()}
        report["protected_after_failure"] = protected_snapshot(source, source_neighbor, output_neighbor)
        require(overlap["failure_stage_absent"] and overlap["success_held_after_failure"] == good_held["held_snapshot"]
                and report["prior_after_failure"] == priors and report["protected_after_failure"] == protected,
                "Failed-run cleanup crossed another run's/input ownership")
        LAST_CASE = {"phase": "failed-run-cleaned-other-stage-held", "last_case": bad_row, "held_rows": [good_held], "report_so_far": report}
        overlap["success_release_ns"] = time.monotonic_ns()
        release(good, nonce)
        good_row = publication({**good_held, **good.finish()}, source, base, generator)
        overlap["success_exit_ns"] = time.monotonic_ns()
        COMPLETED_CASES.append(good_row)
        completed = {row["id"]: row["publication"]["snapshot"] for row in COMPLETED_CASES if row["id"] != "overlap-mid-write"}
        report["completed_before_last_repeat"] = {name: flat_snapshot(Path(item["path"])) for name, item in completed.items()}
        child = Child("repeat-after", source, base, controls / "repeat-after", nonce, work)
        children.append(child)
        LAST_CASE = {"phase": "repeat-after", "report_so_far": report}
        COMPLETED_CASES.append(publication(child.finish(), source, base, generator))
        report["completed_after_last_repeat"] = {name: flat_snapshot(Path(item["path"])) for name, item in completed.items()}
        report.update(cases=list(COMPLETED_CASES), prior_after_all={name: flat_snapshot(Path(item["path"])) for name, item in priors.items()},
                      protected_after_all=protected_snapshot(source, source_neighbor, output_neighbor))
        require(report["prior_after_all"] == priors and report["protected_after_all"] == protected, "Later repeat changed an earlier run/input/neighbor")
        successes = {row["id"]: row for row in COMPLETED_CASES if row["id"] != "overlap-mid-write"}
        samples = work / "samples"
        samples.mkdir()
        final = Path(good_row["result"]["execution"]["final_directory"])
        entries = good_row["result"]["execution"]["outputs"]
        sample_records = []
        for name, original in (("source.pdf", source), ("first.pdf", final / entries[0]["filename"]), ("last.pdf", final / entries[-1]["filename"])):
            sample = samples / name
            sample.write_bytes(original.read_bytes())
            sample_records.append({"path": str(sample), "original": str(original), "sha256": history.file_digest(sample)})
        report["render_samples"] = sample_records
        report["base_directories_before_cleanup"] = sorted(path.name for path in base.iterdir() if path.is_dir())
        snapshots = [successes[name]["publication"]["snapshot"] for name in
                     ("prior-repeat-1", "prior-repeat-2", "overlap-success", "repeat-after")] + [bad_row["failure"]["snapshot"]]
        # Validate all acceptance evidence before authorizing any test cleanup.
        report.update(result="RUN_ISOLATION_REGRESSION_PASSED", success=True, exit_code=0,
                      source_after=runner.source_manifest(), source_unchanged=True,
                      cleanup={"authenticated_targets": snapshots, "removed_paths": [],
                               "remaining_members": sorted(path.name for path in base.iterdir())}, owned_output_cleanup=False)
        validator.validate_isolation_report(report, cleanup_complete=False)
        for observed in snapshots:
            is_failed = Path(observed["path"]).name.startswith(".WinBookSplit-failed-")
            owner = json.loads((Path(observed["path"]) / ".WinBookSplit-owner.json").read_text(encoding="utf-8"))
            require(owner["kind"] == ("failed" if is_failed else "run"), "Owned cleanup kind changed")
            remove_authenticated(observed, owner)
            report["cleanup"]["removed_paths"].append(observed["path"])
        report["cleanup"]["remaining_members"] = sorted(path.name for path in base.iterdir())
        report["protected_after_cleanup"] = protected_snapshot(source, source_neighbor, output_neighbor)
        report["owned_output_cleanup"] = True
        report["source_after"] = runner.source_manifest()
        validator.validate_isolation_report(report)
        LAST_CASE = {"phase": "complete", "report_so_far": report}
        return report
    except BaseException as error:
        primary_error = error
        LAST_CASE["owned_child_observations"] = [child.observation() for child in children]
        raise
    finally:
        recover_children(children, primary_error)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--sample-directory", type=Path)
    args = parser.parse_args()
    target = runner.new_external_path(args.report)
    temporary = tempfile.TemporaryDirectory(prefix="I-")
    work = Path(temporary.name).resolve()
    try:
        report = characterize(work)
        if args.sample_directory is not None:
            samples = args.sample_directory
            require(samples.is_absolute() and not samples.exists() and not samples.resolve().is_relative_to(ROOT), "New external sample directory required")
            samples.mkdir()
            for item in report["render_samples"]:
                destination = samples / Path(item["path"]).name
                destination.write_bytes(Path(item["path"]).read_bytes())
                item["external_path"] = str(destination)
                require(history.file_digest(destination) == item["sha256"], "External render sample differs from actual written PDF")
        validator.validate_isolation_report(report)
        temporary.cleanup()
        report["owned_temp_removed"] = not work.exists()
        with target.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(report, stream, ensure_ascii=True, indent=2)
            stream.write("\n")
        print("Isolation acceptance passed: four complete runs, one real mid-write failure, two simultaneously held stages")
        return 0
    except BaseException as error:
        temporary._finalizer.detach()
        with target.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump({"schema_version": 1, "task_id": "M4-T03", "result": "RUN_ISOLATION_REGRESSION_FAILED",
                       "success": False, "exit_code": 1, "error": str(error), "cleanup_safe": False,
                       "workspace_retained": str(work) if work.exists() else None, "owned_temp_removed": not work.exists(),
                       "completed_cases": COMPLETED_CASES, "last_case": LAST_CASE}, stream, ensure_ascii=True, indent=2)
            stream.write("\n")
        print("Isolation acceptance failed; authored workspace retained " + str(work) + ": " + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
