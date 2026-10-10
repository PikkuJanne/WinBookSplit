"""AC-075 narrow actual-host execution of captured PDFs after owned path faults.

The copied engine adapter changes only an authored processing copy after the
real planner has captured it. It neither replaces the writer nor fabricates a
result. Legacy path/stream/exit/cancel cases run in their existing layers.
"""

from __future__ import annotations

from contextlib import redirect_stdout
import base64
from datetime import datetime, timezone
from hashlib import sha256
import importlib.util
from io import BytesIO, StringIO
import json
import os
from pathlib import Path
import shutil
import sys

from pypdf import PdfReader, PdfWriter

ROOT = Path(__file__).resolve().parents[2]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


process = load("wbs_cross_host_process", ROOT / "tests/process/characterize_process.py")
cli = load("wbs_cross_host_cli", ROOT / "tests/cli/characterize_cli.py")
validator = load("wbs_cross_host_validator", ROOT / "tests/faults/validate_cross_host_report.py")
manual, history, paths, launchers, runner = process.manual, process.history, process.paths, process.launchers, process.runner
need = history.require
KINDS = ("replace", "delete")
LAST_CASE = None
COMPLETED_CASES = []

ADAPTER = r'''"""Declared owned-copy source fault; the saved shipped engine still plans/writes."""
import importlib.util
import json
import os
from pathlib import Path
import sys
from hashlib import sha256

saved = Path(__file__).with_name('saved_engine.py')
spec = importlib.util.spec_from_file_location('wbs_captured_source_engine', saved)
engine = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = engine
spec.loader.exec_module(engine)
source = Path(os.environ['WBS_SNAPSHOT_SOURCE'])
base = Path(os.environ['WBS_SNAPSHOT_BASE'])
target = Path(os.environ['WBS_SNAPSHOT_RECEIPT'])
kind = os.environ['WBS_SNAPSHOT_KIND']
source_key = os.path.normcase(os.path.abspath(source))
reader = engine.PdfReader
prepare = engine.prepare_split
observation = {'protocol':'winbooksplit.authored-source-fault','version':1,'kind':kind,
    'pid':os.getpid(),'parent_pid':os.getppid(),'python':sys.executable,'argv':sys.argv[1:],
    'saved_engine_sha256':sha256(saved.read_bytes()).hexdigest(),
    'source_reader_calls':0,'source_reread_attempts':[],'staged_output_reader_paths':[],
    'state':'starting','prepared_plan':None}

def identity(path):
    if not path.exists():
        return {'path':str(path),'exists':False}
    details = path.lstat()
    data = path.read_bytes()
    return {'path':str(path),'exists':True,'device':details.st_dev,'inode':details.st_ino,
        'attributes':details.st_file_attributes,'size_bytes':len(data),'sha256':sha256(data).hexdigest()}

def save(first=False):
    with target.open('x' if first else 'w', encoding='utf-8', newline='\n') as stream:
        json.dump(observation, stream, ensure_ascii=True, separators=(',',':'))
        stream.write('\n')

def tracked_reader(value, *args, **kwargs):
    if isinstance(value, (str, os.PathLike)):
        path = Path(value)
        key = os.path.normcase(os.path.abspath(path))
        if key == source_key:
            if observation['prepared_plan'] is not None:
                observation['source_reread_attempts'].append(str(path))
                save()
                raise RuntimeError('Authored control forbids reopening the stale processing path')
            observation['source_reader_calls'] += 1
        elif observation['prepared_plan'] is not None:
            if path.parent.parent != base or not path.parent.name.startswith('.WinBookSplit-stage-'):
                raise RuntimeError('Unexpected reader path outside this owned output stage')
            observation['staged_output_reader_paths'].append(str(path))
    return reader(value, *args, **kwargs)

def source_fault(input_path, mode, manual_data=None, **kwargs):
    prepared = prepare(input_path, mode, manual_data, **kwargs)
    if observation['prepared_plan'] is not None or os.path.normcase(os.path.abspath(input_path)) != source_key:
        raise RuntimeError('The authored source control must prepare exactly once')
    observation['prepared_plan'] = json.loads(engine.serialize_result(prepared.plan))
    observation['before_fault'] = identity(source)
    if kind == 'replace':
        replacement = Path(os.environ['WBS_SNAPSHOT_REPLACEMENT']).read_bytes()
        source.write_bytes(replacement)
    elif kind == 'delete':
        source.unlink()
    else:
        raise RuntimeError('Unknown authored source fault')
    observation['after_fault'] = identity(source)
    observation['state'] = 'prepared_source_fault_applied'
    save(first=True)
    return prepared

engine.PdfReader = tracked_reader
engine.prepare_split = source_fault
result = engine.main()
observation['state'] = 'engine_completed'
observation['engine_main_return'] = result
observation['engine_system_exit_code'] = 0 if result is None else result
save()
raise SystemExit(result)
'''


def digest(path):
    return sha256(path.read_bytes()).hexdigest()


def identity(path):
    if not path.exists():
        return {"path": str(path), "exists": False}
    need(path.is_file() and not history.is_reparse(path), "Snapshot input must be an ordinary authored file")
    details = path.lstat()
    return {"path": str(path), "exists": True, "device": details.st_dev, "inode": details.st_ino,
        "attributes": details.st_file_attributes, "size_bytes": details.st_size, "sha256": digest(path)}


def tree_identity(directory):
    need(directory.is_dir() and not history.is_reparse(directory), "Prior run must be an ordinary owned directory")
    details = directory.lstat()
    members = {}
    for member in sorted(directory.iterdir()):
        need(member.is_file() and not history.is_reparse(member), "Prior run has an unexpected directory/reparse member")
        members[member.name] = identity(member)
    return {"path": str(directory), "device": details.st_dev, "inode": details.st_ino, "members": members}


def remove_run(base, execution, authenticated):
    directory = manual.published_directory(base, execution)
    actual = tree_identity(directory)
    expected = {".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", *(item["filename"] for item in execution["outputs"])}
    need(set(actual["members"]) == expected and actual == authenticated,
        "Snapshot cleanup refuses an unexpected or changed authenticated publication member")
    launchers.remove_known_directory(directory, base, expected, ".WinBookSplit-owner.json",
        {"schema_version": 1, "kind": "run", "run_id": execution["run_id"]},
        expected_sha256={name: item["sha256"] for name, item in actual["members"].items()},
        expected_file_identities=actual["members"],
        expected_directory_identity={"device": actual["device"], "inode": actual["inode"]})


def case(work, host, kind, generator, original_bytes, replacement_bytes, engine):
    global LAST_CASE
    label = host["id"] + "-captured-" + kind
    directory = work / label
    directory.mkdir()
    app, books, cwd, base = [directory / name for name in ("a", "b", "c", "o")]
    application_sha = process.copy_application(app)
    for path in (books, cwd, base):
        path.mkdir()
    original, processing, replacement, neighbor = [books / name for name in
        ("immutable-original.pdf", "processing ää 日本 [1] & !%.pdf", "replacement.pdf", "neighbor.txt")]
    for path, data in ((original, original_bytes), (processing, original_bytes), (replacement, replacement_bytes),
            (neighbor, b"Preserve this authored snapshot neighbor\n")):
        with path.open("xb") as stream:
            stream.write(data)
    # A genuine prior publication is a direct shipped-engine API precondition,
    # not an extra native launcher case or a synthesized success manifest.
    prior_stdout = StringIO()
    with redirect_stdout(prior_stdout):
        prior_prepared = engine.prepare_split(original, "manual", "1,3", output_base=base)
        prior_execution = manual.plain(engine.execute_split(prior_prepared, base))
    prior_outputs = manual.published_outputs(base, generator, prior_execution)
    prior_directory = Path(prior_execution["final_directory"])
    source_before = {"original": identity(original), "replacement": identity(replacement), "neighbor": identity(neighbor),
        "prior": tree_identity(prior_directory)}
    processing_before = identity(processing)
    shutil.copyfile(app / "engine/winbooksplit_engine.py", app / "engine/saved_engine.py")
    (app / "engine/winbooksplit_engine.py").write_bytes(ADAPTER.encode("utf-8"))
    copied_after = {name: digest(app / name) for name in process.APPLICATION}
    adapter_path = directory / "source-fault.json"
    parameters = ["-InputFile", str(processing), "-OutputDirectory", str(base), "-PythonPath", sys.executable,
        "-Mode", "Manual", "-StartPages", "1,3", "-NonInteractive", "-NoPause"]
    command = [host["shell_executable"], "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned",
        "-File", str(app / "WinBookSplit.ps1"), *parameters]
    environment = history.clean_environment(cwd)
    environment.update(PSModulePath=str(Path(host["shell_executable"]).parent / "Modules"),
        WBS_SNAPSHOT_SOURCE=str(processing), WBS_SNAPSHOT_BASE=str(base), WBS_SNAPSHOT_RECEIPT=str(adapter_path),
        WBS_SNAPSHOT_KIND=kind, WBS_SNAPSHOT_REPLACEMENT=str(replacement))
    LAST_CASE = {"id": label, "host_id": host["id"], "kind": kind, "passed": False, "command": command,
        "cwd": str(cwd), "source_observations_before": source_before, "processing_copy_before": processing_before,
        "application_sha256": application_sha, "copied_application_sha256": copied_after,
        "adapter_sha256": digest(app / "engine/winbooksplit_engine.py"), "source_fault_receipt_path": str(adapter_path)}
    try:
        with paths.read_only_source(original):
            readonly = bool(original.lstat().st_file_attributes & 1)
            actual = cli.run_unavailable_stdin(command, cwd, environment, timeout=60)
        LAST_CASE.update(actual, actual_process=True, source_fault_receipt=process.owned_artifact(adapter_path),
            processing_copy_after=identity(processing))
        final = cli.outcome(actual["stdout"])
        log = launchers.exact_line(actual["stdout"], "Log: ")
        need(log.parent.parent == base and log.name == "console.log", "Snapshot console escaped its explicit output base")
        raw_log = log.read_bytes()
        text = raw_log.decode("utf-8", errors="strict")
        fault_raw = adapter_path.read_bytes()
        fault = json.loads(fault_raw)
        frame = final["engine_result"]
        execution = frame["execution"]
        outputs = manual.published_outputs(base, generator, execution)
        run = Path(execution["final_directory"])
        publication = tree_identity(run)
        content = [value for item in outputs for value in process.conversion.page_content(PdfReader(run / item["filename"]))]
        source_after = {"original": identity(original), "replacement": identity(replacement), "neighbor": identity(neighbor),
            "prior": tree_identity(prior_directory)}
        records = [json.loads(line) for line in text.splitlines() if line.startswith('{')]
        summaries = [json.loads(line[len("[PROCESS] "):]) for line in text.splitlines() if line.startswith("[PROCESS] ")]
        invocations = [json.loads(line[len("[ENGINE] "):]) for line in text.splitlines() if line.startswith("[ENGINE] ")]
        console = launchers.authenticate_console_manifest(log, final, source_path=processing,
            source_observation={"sha256": processing_before["sha256"], "size_bytes": processing_before["size_bytes"]})
        record = {**LAST_CASE, "passed": True, "parameters": parameters, "python_executable": sys.executable,
            "shell_executable": host["shell_executable"], "host_version": host["host_version"], "output_base": str(base),
            "processing_copy_after": identity(processing), "source_observations_after": source_after,
            "source_readonly_observed": readonly, "source_fault": fault, "source_fault_raw": fault_raw.decode("utf-8"),
            "source_fault_sha256": sha256(fault_raw).hexdigest(), "saved_engine_sha256": digest(app / "engine/saved_engine.py"),
            "outcome": final, "engine_records": records, "process_summaries": summaries, "engine_invocations": invocations,
            "console_log": text, "console_log_sha256": sha256(raw_log).hexdigest(), "console_evidence": console,
            "outputs": outputs, "source_page_ids": generator.page_ids(original),
            "publication_observation": publication,
            "publication_owner_raw": (run / ".WinBookSplit-owner.json").read_bytes().decode("utf-8"),
            "publication_manifest_raw": (run / "WinBookSplit_Manifest.json").read_bytes().decode("utf-8"),
            "source_content_sha256": process.conversion.page_content(PdfReader(original)), "output_content_sha256": content,
            "prior_setup": {"actual_engine_api": True, "native_launcher_case": False, "execution": prior_execution,
                "outputs": prior_outputs, "stdout": prior_stdout.getvalue()}, "owned_outputs_removed": False,
            "source_fault_scope": "Only the authored processing copy is changed by the declared adapter after capture"}
        LAST_CASE = {**record, "passed": False}
        validator.validate_case(record, host, runner.source_manifest(), cleanup_complete=False)
        remove_run(base, execution, publication)
        launchers.remove_console_directory(log, base, final, source_path=processing,
            source_observation={"sha256": processing_before["sha256"], "size_bytes": processing_before["size_bytes"]})
        need(tree_identity(prior_directory) == source_before["prior"], "Own cleanup reached the prior publication")
        record.update(owned_outputs_removed=True, marked_stage_absent=not any(path.name.startswith('.WinBookSplit-stage-') for path in base.iterdir()))
        remove_run(base, prior_execution, source_before["prior"])
        need(not list(base.iterdir()), "Snapshot cleanup left unexpected output-base members")
        record["source_observations_after_cleanup"] = {name: identity(path) for name, path in
            (("original", original), ("replacement", replacement), ("neighbor", neighbor))}
        record["processing_copy_after_cleanup"] = identity(processing)
        record["prior_removed_after_preservation_proof"] = True
        validator.validate_case(record, host, runner.source_manifest(), cleanup_complete=True)
        return record
    except BaseException as error:
        # Parent driver retains its workspace. Only successful authenticated
        # records authorize exact marked output cleanup above.
        if adapter_path.is_file():
            LAST_CASE["source_fault_receipt"] = process.owned_artifact(adapter_path)
        LAST_CASE["processing_copy_after"] = identity(processing)
        if isinstance(error, history.EntryPointFailure):
            observation = {"error_type": type(error).__name__, "message": str(error), "tracked_parent_pid": error.pid,
                "native_exit_code": None, "raw_streams_complete": False, "inner_tree_shutdown_proved": False}
            cause = error.__cause__
            for field in ("stdout", "stderr"):
                raw = getattr(cause, field, None)
                if type(raw) is bytes and len(raw) <= 16 * 1024 * 1024:
                    observation[field + "_partial_base64"] = base64.b64encode(raw).decode("ascii")
                    observation[field + "_partial_sha256"] = sha256(raw).hexdigest()
                    observation[field + "_partial_size_bytes"] = len(raw)
            LAST_CASE["supervision_error"] = observation
        raise


def characterize_cross_host(work, shells):
    global LAST_CASE
    LAST_CASE = None
    COMPLETED_CASES.clear()
    need(os.name == "nt" and len(shells) == 2, "Both actual supported Windows hosts required")
    before = runner.source_manifest()
    work = Path(work)
    work.mkdir()
    root = work / "cross-host"
    root.mkdir()
    probe = process.host_probe(root)
    hosts = [paths.host_observation(Path(shell), root, probe) for shell in shells]
    need({host["id"] for host in hosts} == {"PS51", "PS7"}, "Cross-host cases require actual distinct supported hosts")
    generator = history.load_generator()
    original = generator._pdf_bytes({"title": "Original generated M4-T03 captured-source fixture", "pages": 4, "outline": []})
    writer = PdfWriter()
    for page in reversed(PdfReader(BytesIO(original)).pages):
        writer.add_page(page)
    buffer = BytesIO()
    writer.write(buffer)
    replacement = buffer.getvalue()
    need(original != replacement, "Authored reverse-page replacement must differ")
    engine = load("wbs_cross_host_prior_engine", ROOT / "engine/winbooksplit_engine.py")
    for host in hosts:
        for kind in KINDS:
            COMPLETED_CASES.append(case(root, host, kind, generator, original, replacement, engine))
    cases = list(COMPLETED_CASES)
    for host in hosts:
        after = paths.host_observation(Path(host["shell_executable"]), root, probe)
        host["policies_after"] = after["stored_policies"]
    report = {"protocol": "winbooksplit.fault-cross-host", "version": 1, "task_id": "M4-T03", "acceptance_ids": ["AC-075"],
        "result": "CROSS_HOST_SNAPSHOT_PASSED", "success": True, "exit_code": 0, "case_count": 4,
        "host_cases": hosts, "cases": cases, "tested_path_sha256": before,
        "source_after_sha256": runner.source_manifest(), "source_unchanged": before == runner.source_manifest(),
        "fixture_workspace_owned_by_parent": True, "owned_outputs_removed": True, "machine_settings_unchanged": True,
        "cleanup_scope": "Only the application's marked console/chapter/prior outputs; parent fault driver owns fixture workspace",
        "observed_at": datetime.now(timezone.utc).isoformat(),
        "environment": {"python": sys.version, "python_executable": sys.executable, "pypdf": history.pypdf.__version__},
        "not_run": ["human or GUI testing", "independent legacy path/process/outcome suites inside this section"],
        "limits": ["Four copied-engine adapter cases exercise the real saved planner and writer; source mutation is an authored control.",
            "AC-075 additionally requires same-source actual existing paths/process/outcomes receipts."]}
    validator.validate_cross_host_report(report, [str(Path(shell).resolve()) for shell in shells], cleanup_complete=False)
    return report
