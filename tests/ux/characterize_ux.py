"""M3-T04 native held-plan confirmation and real converted-PDF lifetime controls."""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
from hashlib import sha256
import importlib.util
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import threading
import time

from pypdf import PdfReader


ROOT = Path(__file__).resolve().parents[2]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


cli = load("wbs_ux_cli", ROOT / "tests/cli/characterize_cli.py")
validator = load("wbs_ux_validator", ROOT / "tests/ux/validate_ux_report.py")
process_tests, history, paths, conversion, launchers, runner, require = \
    cli.process_tests, cli.history, cli.paths, cli.conversion, cli.launchers, cli.runner, cli.require
COMPLETED_CASES, LAST_CASE = [], None
ACTIVE_DIALOGUES = []
CONFIRM = "Create these chapter PDFs? [Y] yes, [C] cancel"
MANUAL = "Pages (comma separated; blank or C cancels)"


class NativeDialogue:
    """Concurrent bounded observations of one actual retained process handle."""
    def __init__(self, command, cwd, environment, timeout=90, unavailable=False):
        self.command, self.cwd, self.timeout = command, cwd, timeout
        self.started = datetime.now(timezone.utc).isoformat()
        self.clock = time.monotonic()
        self.child = subprocess.Popen(command, cwd=cwd, env=environment,
            stdin=subprocess.DEVNULL if unavailable else subprocess.PIPE,
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, bufsize=0)
        self.lock, self.buffers, self.errors, self.eof = threading.Lock(), {"stdout": bytearray(), "stderr": bytearray()}, [], set()
        self.answers, self.threads = [], []
        for name in self.buffers:
            thread = threading.Thread(target=self.drain, args=(name,), daemon=True)
            thread.start()
            self.threads.append(thread)
        ACTIVE_DIALOGUES.append(self)

    def drain(self, name):
        try:
            while True:
                block = os.read(getattr(self.child, name).fileno(), 8192)
                if not block:
                    break
                with self.lock:
                    if len(self.buffers[name]) + len(block) > 8388608:
                        raise ValueError("Native UX stream exceeded its 8 MiB capture bound")
                    self.buffers[name].extend(block)
            with self.lock:
                self.eof.add(name)
        except BaseException as error:
            with self.lock:
                self.errors.append(str(error))

    def text(self, name="stdout", final=False):
        with self.lock:
            return bytes(self.buffers[name]).decode("utf-8", "strict" if final else "replace")

    def wait_text(self, text):
        while text not in self.text():
            if self.errors or self.child.poll() is not None or time.monotonic() - self.clock >= self.timeout:
                self.fail("Native UX process did not reach the expected prompt: " + text)
            time.sleep(.01)
        return self.text()

    def send(self, text):
        require(self.child.poll() is None and self.child.stdin is not None, "Cannot answer an exited/closed UX process")
        raw = text.encode("utf-8")
        self.child.stdin.write(raw)
        self.child.stdin.flush()
        self.answers.append(text)

    def close_input(self):
        if self.child.stdin is not None:
            self.child.stdin.close()
            self.child.stdin = None

    def fail(self, message):
        if self.child.poll() is None:
            self.child.terminate()  # Only this retained parent handle; descendants remain unproved.
        raise history.EntryPointFailure(message + "; descendant shutdown/cleanup unproved, retain workspace",
                                        cleanup_safe=False, pid=self.child.pid)

    def finish(self, close_input=True):
        if close_input:
            self.close_input()
        try:
            self.child.wait(timeout=max(.1, self.timeout - (time.monotonic() - self.clock)))
        except subprocess.TimeoutExpired:
            self.fail("Native UX process exceeded bounded deadline")
        for thread in self.threads:
            thread.join(2)
        if self.errors or self.eof != {"stdout", "stderr"}:
            self.fail("Native UX stream EOF/error proof incomplete: " + repr(self.errors))
        self.close_input()
        result = {"command": self.command, "cwd": str(self.cwd), "started_at": self.started,
                  "finished_at": datetime.now(timezone.utc).isoformat(), "pid": self.child.pid,
                  "exit_code": self.child.returncode, "elapsed_seconds": time.monotonic() - self.clock,
                  "timeout_seconds": self.timeout, "timed_out": False, "streams_complete": True,
                  "stdin_utf8": "".join(self.answers)}
        for name in self.buffers:
            result[name] = self.text(name, final=True)
            result[name + "_sha256"] = sha256(result[name].encode("utf-8")).hexdigest()
            getattr(self.child, name).close()
        ACTIVE_DIALOGUES.remove(self)
        return result


def held_without_chapters(native, base):
    require(native.child.poll() is None, "Confirmation did not retain the actual process")
    members = sorted(item.name for item in base.iterdir())
    console = [item for item in base.iterdir() if item.name.startswith(".WinBookSplit-console-")]
    require(len(console) == 1 and set(members) == {"prior-output.pdf", console[0].name},
            "Chapters or a stage were reserved before consent")
    names = sorted(item.name for item in console[0].iterdir())
    require(names == [".WinBookSplit-console-owner.json", "console.log"], "Held console ownership/membership changed")
    marker = json.loads((console[0] / names[0]).read_text(encoding="utf-8"))
    require(marker == {"run_id": console[0].name.removeprefix(".WinBookSplit-console-"), "kind": "console"},
            "Held console marker is not this run's owner")
    return {"parent_running": True, "base_members": members, "console_members": names,
            "console_owner": marker, "chapter_or_stage_absent": True, "stdout": native.text()}


def publication(base, execution, expected_content):
    final = Path(execution["final_directory"])
    names = {".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", *[item["filename"] for item in execution["outputs"]]}
    require(final.parent == base and {item.name for item in final.iterdir()} == names,
            "Actual UX publication has unexpected membership")
    manifest = json.loads((final / "WinBookSplit_Manifest.json").read_text(encoding="utf-8"))
    require(manifest == execution["manifest"], "Actual UX manifest differs from execution")
    observed = []
    for item in execution["outputs"]:
        target = final / item["filename"]
        content = conversion.page_content(PdfReader(target))
        require(content == expected_content[item["start"]:item["end"]] and
                len(content) == item["page_count"] and cli.digest(target) == item["sha256"] and
                target.stat().st_size == item["size_bytes"], "Actual UX chapter bytes/pages differ from immutable plan")
        observed.extend(content)
    require(observed == expected_content, "Actual UX omitted/reordered/duplicated physical pages")
    result = {"manifest": manifest, "manifest_raw": (final / "WinBookSplit_Manifest.json").read_bytes().decode("utf-8"),
              "manifest_sha256": cli.digest(final / "WinBookSplit_Manifest.json"),
              "members": sorted(names), "content_sha256": observed}
    launchers.remove_known_directory(final, base, names, ".WinBookSplit-owner.json",
                                    {"schema_version": 1, "kind": "run", "run_id": execution["run_id"]})
    return result


def application_case(work, host, kind, generator, ebooks, references, calibre):
    global LAST_CASE
    directory = work / (host["id"] + "-" + kind)
    directory.mkdir()
    app, books, cwd, base = [directory / name for name in ("a", "b", "c", "o")]
    copied = process_tests.copy_application(app)
    for item in (books, cwd, base):
        item.mkdir()
    fmt = "epub" if kind == "ebook-manual-epub" else "azw3" if kind == "ebook-manual-azw3" else "pdf"
    source = books / ("Authored Å 日本 [1] & %! $(literal)." + fmt)
    if fmt != "pdf":
        shutil.copyfile(ebooks[fmt], source)
    else:
        outline = [] if kind.startswith("fallback-") else [
            {"title": "Parent A Å", "start_page": 3, "children": [
                {"title": "Child A1", "start_page": 4, "children": []},
                {"title": "Child A2", "start_page": 7, "children": []}]},
            {"title": "Parent B childless", "start_page": 9, "children": []}]
        source.write_bytes(generator._pdf_bytes({"title": "Authored complete twelve physical pages", "pages": 12, "outline": outline}))
    neighbor = books / "authored-neighbor.txt"
    neighbor.write_bytes(b"Authored neighboring input stays unchanged\n")
    prior = base / "prior-output.pdf"
    prior.write_bytes(b"Authored earlier output stays unchanged\n")
    environment = history.clean_environment(cwd)
    system = Path(os.environ["SystemRoot"])
    environment.update(PATH=os.pathsep.join((str(Path(host["shell_executable"]).parent), str(system / "System32"), str(system))),
                       PSModulePath=str(Path(host["shell_executable"]).parent / "Modules"))
    parameters = ["-InputFile", str(source), "-OutputDirectory", str(base), "-PythonPath", sys.executable]
    manual = kind in {"manual-physical", "ebook-manual-epub", "ebook-manual-azw3"}
    if not manual:
        parameters += ["-Mode", "Auto", "-BookmarkLevel", "1" if kind.startswith("fallback-") else "2"]
    if fmt != "pdf":
        parameters += ["-CalibrePath", str(calibre)]
    if kind == "noninteractive":
        parameters += ["-NonInteractive"]
    elif kind == "preview":
        parameters += ["-Preview"]
    elif kind == "pending-timeout":
        parameters += ["-ProcessTimeout", "3"]
    else:
        parameters += ["-NoPause"]
    command = [host["shell_executable"], "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", str(app / "WinBookSplit.ps1"), *parameters]
    unavailable = kind in {"noninteractive", "preview"}
    if unavailable:
        command.insert(2, "-NonInteractive")
    LAST_CASE = {"id": host["id"] + "-" + kind, "command": command, "workspace": str(directory)}
    with paths.read_only_source(source):
        before = {name: cli.identity(item) for name, item in (("source", source), ("neighbor", neighbor), ("prior", prior))}
        native = NativeDialogue(command, cwd, environment, unavailable=unavailable)
        hold = None
        if manual:
            native.send("M\n")
            if fmt != "pdf":
                native.wait_text("Open the generated PDF?")
                native.send("N\n")
            native.wait_text(MANUAL)
            native.send("1,2,3\n" if fmt != "pdf" else "1,5,9\n")
        if kind == "fallback-cancel":
            native.wait_text("Choose [M] manual or [C] cancel")
            native.send("C\n")
        elif kind == "fallback-manual":
            native.wait_text("Choose [M] manual or [C] cancel")
            native.send("M\n")
            native.wait_text(MANUAL)
            native.send("1,5,9\n")
        if not unavailable and kind != "fallback-cancel":
            native.wait_text(CONFIRM)
            hold = held_without_chapters(native, base)
            if kind == "confirm-cancel":
                native.send("C\n")
            elif kind == "confirm-blank":
                native.send("\n")
            elif kind == "confirm-eof":
                native.close_input()
            elif kind == "pending-timeout":
                pass  # Keep input open and send no future answer.
            else:
                native.send("Q\nmaybe\nY\nN\n" if kind == "confirm-invalid" else "Y\nN\n")
        observed = native.finish(close_input=kind != "pending-timeout")
        # Preserve completed raw streams even if later evidence parsing fails.
        LAST_CASE.update(observed)
        LAST_CASE.update(application_sha256=copied, source_observations_before=before,
                         source_observations_after={name: cli.identity(item) for name, item in (("source", source), ("neighbor", neighbor), ("prior", prior))},
                         output_members_observed=sorted(item.name for item in base.iterdir()))
        final = cli.outcome(observed["stdout"])
        LAST_CASE["outcome"] = final
        log, results, transports, log_hash = cli.console_records(observed["stdout"], base)
        if kind == "preview":
            require(log is None and results == transports == [] and log_hash is None,
                    "Preview unexpectedly persisted console/transport records")
            text = None  # Preview deliberately uses an unpersisted StringWriter.
        else:
            require(log is not None and len(transports) == 1, "Actual UX console/transport records missing")
            text = log.read_bytes().decode("utf-8", "strict")
        lines = (text or "").splitlines()
        plans = [json.loads(line[len("[PLAN] "):]) for line in lines if line.startswith("[PLAN] ")]
        replies = [json.loads(line[len("[INTERACTION-REPLY] "):]) for line in lines if line.startswith("[INTERACTION-REPLY] ")]
        interactions = [json.loads(line[len("[WBS-INTERACTION] "):]) for line in lines if line.startswith("[WBS-INTERACTION] ")]
        content = conversion.page_content(PdfReader(source if fmt == "pdf" else references[fmt]["path"]))
        published = publication(base, final["engine_result"]["execution"], content) if final["status"] == "success" else None
        if final["engine_result"]:
            conversion.remove_diagnostic(base, final["engine_result"])
        row = {"id": host["id"] + "-" + kind, "kind": kind, "host_id": host["id"],
               "host_version": host["host_version"], "shell_executable": host["shell_executable"],
               "application_sha256": copied, "parameters": parameters, "actual_process": True,
               "input_path": str(source), "output_base": str(base), "source_read_only_observed": True,
               "source_observations_before": before, "source_observations_after": {name: cli.identity(item) for name, item in (("source", source), ("neighbor", neighbor), ("prior", prior))},
               "hold": hold, "outcome": final, "engine_records": results, "process_summaries": transports,
               "console_log": text, "console_log_sha256": log_hash, "plan_events": plans, "interaction_replies": replies,
               "terminal_evidence_source": "stdout_outcome" if kind == "preview" else "console_log",
               "interaction_events": interactions,
               "expected_content_sha256": content, "publication": published, **observed}
        LAST_CASE = row
        if log is not None:
            marker = {"run_id": log.parent.name.removeprefix(".WinBookSplit-console-"), "kind": "console"}
            launchers.remove_known_directory(log.parent, base, {".WinBookSplit-console-owner.json", "console.log"}, ".WinBookSplit-console-owner.json", marker)
        row.update(source_observations_after_cleanup={name: cli.identity(item) for name, item in (("source", source), ("neighbor", neighbor), ("prior", prior))},
                   application_unchanged=copied == {name: cli.digest(app / name) for name in copied},
                   output_members_after_cleanup=sorted(item.name for item in base.iterdir()), passed=True)
        require(row["source_observations_before"] == row["source_observations_after"] == row["source_observations_after_cleanup"], "UX changed source/neighbor/prior")
        validator.validate_application_case(row)
        COMPLETED_CASES.append(row)
        LAST_CASE = row
        return row


def working_pdf_case(work, fmt, source, reference, calibre):
    directory = work / ("engine-working-" + fmt)
    directory.mkdir()
    base = directory / "out"
    base.mkdir()
    copied_source = directory / ("Authored [1] Å & %." + fmt)
    shutil.copyfile(source, copied_source)
    prior = base / "prior-output.pdf"
    prior.write_bytes(b"Authored prior stays unchanged\n")
    session = "b" * 32
    command = [sys.executable, "-I", "-B", "-X", "utf8", str(ROOT / "engine/winbooksplit_engine.py"),
               str(copied_source), str(base), "manual", "", "--interactive", session, "--calibre-path", str(calibre)]
    native = NativeDialogue(command, directory, history.clean_environment(directory))
    events, known_lines = [], 0
    def next_event():
        nonlocal known_lines
        while True:
            text = native.text()
            lines = text.splitlines()
            if text and not text.endswith("\n"):
                lines = lines[:-1]
            for line in lines[known_lines:]:
                if line.startswith("[WBS-INTERACTION] "):
                    value = json.loads(line[len("[WBS-INTERACTION] "):])
                    if not events or value["sequence"] > events[-1]["sequence"]:
                        events.append(value)
                        return value
            if native.child.poll() is not None or time.monotonic() - native.clock > native.timeout:
                native.fail("Direct working PDF event absent")
            time.sleep(.01)
    def reply(event, action, **data):
        native.send(json.dumps({"protocol": "winbooksplit.interaction", "version": 1, "session": session,
                               "sequence": event["sequence"], "action": action, **data}, separators=(",", ":")) + "\n")
    source_before = cli.identity(copied_source)
    event = next_event()
    require(event["stage"] == "input_ready", "Actual conversion did not precede manual starts")
    require(event["conversion"]["workspace_cleanup"]["cleanup_complete"] and not Path(event["source_identity"]["path"]).exists(), "Original conversion workspace was not safely cleaned")
    reply(event, "request_working_pdf")
    event = next_event()
    require(event["stage"] == "working_pdf_ready", "Requested working PDF was not made available")
    target = Path(event["working_pdf"]["path"])
    first = cli.identity(target)
    content = conversion.page_content(PdfReader(target))
    require(content == reference["page_content_sha256"] and first["sha256"] == event["source_identity"]["sha256"], "Actual working PDF differs from captured conversion")
    require(native.child.poll() is None, "Working PDF selection process did not remain alive")
    reply(event, "starts", starts="1,2,3")
    event = next_event()
    require(event["stage"] == "plan_ready" and cli.identity(target) == first, "Working PDF did not remain unchanged through manual selection")
    reply(event, "cancel")
    observed = native.finish()
    records = [json.loads(line) for line in observed["stdout"].splitlines() if line.startswith("{")]
    require(len(records) == 1, "Direct working PDF did not emit one terminal result")
    row = {"id": "engine-working-" + fmt, "format": fmt, "actual_process": True, "viewer_opened": False,
           "events": events, "working_identity": first, "working_content_sha256": content,
           "reference_content_sha256": reference["page_content_sha256"], "working_readable_before_starts": True,
           "working_identity_unchanged_at_plan": True, "working_pdf_absent_after": not target.exists(),
           "working_directory_absent_after": not target.parent.exists(), "source_before": source_before,
           "source_after": cli.identity(copied_source), "output_members_after": sorted(item.name for item in base.iterdir()),
           "terminal_result": records[0], **observed}
    validator.validate_working_case(row)
    return row


def characterize(work, shells, calibre):
    require(os.name == "nt" and len(shells) == 2, "Actual Windows and both supported hosts required")
    before = runner.source_manifest()
    hosts = [paths.host_observation(shell, work, process_tests.host_probe(work)) for shell in shells]
    require({host["id"] for host in hosts} == {"PS51", "PS7"}, "Both actual distinct hosts required")
    ebooks = work / "ebooks"
    fixtures = conversion.fixtures.generate(ebooks, calibre)
    sources = {fmt: ebooks / ("original-three-chapters." + fmt) for fmt in ("epub", "azw3")}
    references = {fmt: conversion.reference_pdf(source, work / ("reference-" + fmt + ".pdf"), str(calibre), work) for fmt, source in sources.items()}
    cases = [application_case(work, host, kind, history.load_generator(), sources, references, calibre)
             for host in hosts for kind in validator.APPLICATION_KINDS]
    working = [working_pdf_case(work, fmt, sources[fmt], references[fmt], calibre) for fmt in sources]
    for host in hosts:
        host["policies_after"] = paths.host_observation(Path(host["shell_executable"]), work, process_tests.host_probe(work))["stored_policies"]
    require(before == runner.source_manifest(), "Sources changed during UX native acceptance")
    return {"schema_version": 1, "task_id": "M3-T04", "result": "UX_REGRESSION_PASSED", "success": True, "exit_code": 0,
            "observed_at": datetime.now(timezone.utc).isoformat(), "acceptance_ids": ["AC-060", "AC-061", "AC-062"],
            "tested_path_sha256": before, "source_unchanged": True, "host_cases": hosts, "cases": cases,
            "working_pdf_cases": working, "case_count": len(cases) + len(working), "fixture_provenance": fixtures,
            "real_conversion_references": references, "human_or_viewer_opening_tested": False,
            "environment": {"python": sys.version, "python_executable": sys.executable, "pypdf": cli.history.pypdf.__version__},
            "limits": ["Native redirected console controls, not human or Explorer/PDF viewer opening.",
                       "Real working PDF availability uses direct engine protocol; the GUI opt-in is a separate actual check."]}


def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--shell-path", required=True, action="append", type=Path)
    parser.add_argument("--calibre-path", required=True, type=Path)
    args = parser.parse_args()
    target = runner.new_external_path(args.report)
    temporary = tempfile.TemporaryDirectory(prefix="U-")
    work = Path(temporary.name).resolve()
    try:
        require(sys.flags.isolated and sys.dont_write_bytecode, "Use isolated developer Python -I -B")
        report = characterize(work, [path.resolve() for path in args.shell_path], args.calibre_path.resolve())
        temporary.cleanup()
        report["owned_temp_removed"] = not work.exists()
        validator.validate_ux_report(report, [str(path.resolve()) for path in args.shell_path], str(args.calibre_path.resolve()))
        with target.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(report, stream, indent=2, ensure_ascii=True)
        print("UX acceptance passed: " + str(report["case_count"]) + " actual native controls")
        return 0
    except BaseException as error:
        temporary._finalizer.detach()
        stopped = []
        for native in ACTIVE_DIALOGUES:
            stop_error = None
            try:
                if native.child.poll() is None:
                    native.child.terminate()  # Exact retained parent only; never a name/PID scan.
                native.child.wait(timeout=5)
            except BaseException as stop_failure:
                stop_error = str(stop_failure)
            for thread in native.threads:
                thread.join(1)
            stopped.append({"pid": native.child.pid, "command": native.command, "parent_exit_code": native.child.poll(),
                            "parent_stopped": native.child.poll() is not None, "descendants_stopped": None,
                            "stop_error": stop_error, "stdout_observed": native.text(), "stderr_observed": native.text("stderr"),
                            "stream_scope": "bounded available tails; incomplete UTF-8 may use replacement; no EOF/cleanup proof"})
        with target.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump({"schema_version": 1, "task_id": "M3-T04", "result": "UX_REGRESSION_FAILED", "success": False,
                       "exit_code": 1, "error": str(error), "workspace_retained": str(work),
                       "cleanup_safe": False, "completed_cases": COMPLETED_CASES, "last_case": LAST_CASE,
                       "owned_parent_shutdown_attempts": stopped}, stream, indent=2)
        print("UX acceptance failed; retained authored workspace " + str(work) + ": " + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
