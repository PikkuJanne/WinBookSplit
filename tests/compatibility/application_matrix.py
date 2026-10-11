"""Native application-runtime cases for the existing shared test runner.

The driver owns fixture generation and PDF oracles. Every application/fixture
child is supervised by the existing Windows job transport and Process probe.
This module neither installs dependencies nor changes application acceptance.
"""

from __future__ import annotations

import base64
from collections.abc import Mapping
from contextlib import ExitStack
from datetime import datetime, timezone
import hashlib
import importlib.util
from importlib import metadata
import json
import os
from pathlib import Path
import re
import stat
import sys
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[2]
UNIT_MODULES = (
    "test_compatibility", "test_filenames", "test_interactive", "test_pdf_policy",
    "test_pdf_fidelity", "test_conversion", "test_outcomes", "test_version", "test_warnings",
)
RESULT_LIMIT = 8 * 1024 * 1024


def require(value, message):
    if not value:
        raise ValueError(message)


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def plain(value):
    if isinstance(value, Mapping):
        return {key: plain(item) for key, item in value.items()}
    if isinstance(value, (tuple, list)):
        return [plain(item) for item in value]
    return value


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def ordinary_file(path):
    """Reject the supplied alias and every reparse ancestor before resolving."""
    path = Path(path)
    require(path.is_absolute(), "A native executable needs an absolute literal path")
    for component in (*reversed(path.parents), path):
        observed = component.lstat()
        require(not observed.st_file_attributes & stat.FILE_ATTRIBUTE_REPARSE_POINT,
                "Native executable paths cannot contain reparse entries")
    require(stat.S_ISREG(path.lstat().st_mode) and path.resolve(strict=True) == path,
            "Native executable must be an ordinary canonical file")
    return path


def validate_process(record):
    """A native exit alone never authorizes evidence-workspace cleanup."""
    require(type(record) is dict and type(record.get("ExitCode")) is int
            and type(record.get("Pid")) is int and record["Pid"] > 0,
            "Native process identity/exit fields are malformed")
    for key in ("ParentStopped", "DescendantsStopped", "StreamsComplete", "JobAssigned", "InputWriterStopped"):
        require(record.get(key) is True, "Unproved native process stop/drain/job: " + key)
    for key in ("TimedOut", "Cancelled", "StdoutTruncated", "StderrTruncated"):
        require(record.get(key) is False, "Native process was interrupted or truncated: " + key)
    for key in ("StartError", "StopError", "StreamError", "ResultError", "InteractionError", "InputError"):
        require(record.get(key) is None, "Native process transport error: " + key)
    for stream in ("Stdout", "Stderr"):
        require(type(record.get(stream)) is str and type(record.get(stream + "TotalBytes")) is int
                and record[stream + "TotalBytes"] == len(record[stream].encode("utf-8")), "Malformed native stream fields/counts")
    return record


def validate_candidate(record, inventory, candidate, policy, policy_sha, compatibility):
    accepted = compatibility.validate_interpreter(record, policy)
    require(accepted["ok"] is True, "Application interpreter rejected: " + accepted["reason"])
    require(record.get("protocol") == "winbooksplit.runtime" and type(record.get("schema_version")) is int
            and record["schema_version"] == 1 and record.get("ok") is True and record.get("reason") == "supported"
            and type(record.get("policy_version")) is int and record["policy_version"] == policy["policy_version"]
            and record.get("policy_sha256") == policy_sha, "Candidate is not bound to the actual shared policy")
    require(Path(record["executable"]) == candidate and inventory.get("executable") == record["executable"]
            and inventory.get("version") == record["version"] and inventory.get("isolated") is True
            and inventory.get("dont_write_bytecode") is True, "Candidate and inventory identity differ")
    require(record.get("pypdf_version") == policy["pypdf"]["version"], "Candidate pypdf differs from the package lock")
    origin = ordinary_file(Path(record.get("pypdf_path", "")))
    roots = [Path(inventory[name]) for name in ("purelib", "platlib")]
    require(any(origin.is_relative_to(root) for root in roots), "Candidate package origin escaped its own library")
    distributions = inventory.get("distributions")
    require(type(distributions) is list and distributions and all(type(item) is dict
            and set(item) == {"name", "version"} and type(item["name"]) is str and type(item["version"]) is str
            for item in distributions), "Malformed runtime-only package inventory")
    names = [item["name"].casefold().replace("_", "-") for item in distributions]
    require(len(names) == len(set(names)) and set(names) <= {"pip", "pypdf"} and "pypdf" in names,
            "This test route requires a runtime-only environment containing pypdf and optional pip")
    require(next(item["version"] for item in distributions if item["name"].casefold() == "pypdf")
            == policy["pypdf"]["version"], "Runtime-only distribution version differs")
    return record


def validate_units(unit, candidate, candidate_version):
    require(type(unit) is dict and unit.get("protocol") == "winbooksplit.application-units"
            and type(unit.get("version")) is int and unit["version"] == 1
            and unit.get("success") is True and unit.get("required_modules") == list(UNIT_MODULES)
            and type(unit.get("modules")) is list and len(unit["modules"]) == len(UNIT_MODULES)
            and all(type(item) is dict and set(item) == {"name", "test_count"} for item in unit["modules"])
            and [item["name"] for item in unit["modules"]] == list(UNIT_MODULES)
            and all(type(item["test_count"]) is int and item["test_count"] > 0 for item in unit["modules"])
            and type(unit.get("tests_run")) is int and unit["tests_run"] == sum(item["test_count"] for item in unit["modules"])
            and all(type(unit.get(key)) is int and unit[key] == 0 for key in
                    ("failures", "errors", "expected_failures", "unexpected_successes"))
            and unit.get("skipped") == [] and unit.get("loading_error") is None
            and unit.get("python_executable") == str(candidate) and unit.get("python_version") == candidate_version
            and unit.get("isolated") is True and unit.get("no_bytecode") is True, "Runtime-only unit receipt differs")
    return unit


def validate_wrapper_runtime(log_text, candidate, full_probe, policy_sha, runtime_receipt, json_reader):
    lines = [line[len("[DEPENDENCY] "):] for line in log_text.splitlines() if line.startswith("[DEPENDENCY] ")]
    require(len(lines) == 1, "Exactly one authenticated wrapper dependency frame is required")
    frame = json_reader(lines[0])
    require(type(frame) is dict and set(frame) == {"Runtime", "Converter"} and frame["Converter"] is None,
            "PDF wrapper dependency frame differs")
    selected = runtime_receipt.runtime(frame["Runtime"])
    require(selected["Source"] == "explicit" and Path(selected["Path"]) == candidate
            and selected["Version"] == full_probe["version"] and selected["PolicySha256"] == policy_sha
            and runtime_receipt.same_json(selected["Details"], full_probe),
            "Wrapper selected runtime differs from the literal candidate/full policy probe")
    return {"raw_frame": lines[0], "frame": frame, "selected_runtime": selected}


class NativeCases:
    """One case per existing PS probe; its outer host also belongs to an owned job."""

    def __init__(self, work, app, shell, supervisor, json_reader):
        self.work, self.app, self.shell = work, app, shell
        self.supervisor, self.json_reader = supervisor, json_reader
        self.records = []
        self.timed_out = False

    def probe(self, case_id, executable, arguments, cwd, *, timeout=120):
        index = len(self.records)
        payload_path, report_path = self.work / f"probe-{index:03}.json", self.work / f"native-{index:03}.json"
        payload = {"helper": str(self.app / "engine/WinBookSplit.Process.ps1"),
                   "diagnostics": str(self.app / "engine/WinBookSplit.Diagnostics.ps1"),
                   "python": str(executable), "work": str(cwd), "cases": [
                       {"id": case_id, "kind": "application-runtime", "arguments": list(map(str, arguments)),
                        "timeout": timeout, "result_limit": RESULT_LIMIT, "cancel": False, "protocol": False}]}
        payload_path.write_text(json.dumps(payload, ensure_ascii=True), encoding="utf-8")
        argv = [str(self.shell), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-File",
                str(ROOT / "tests/process/Probe-Process.ps1"), "-PayloadPath", str(payload_path), "-ReportPath", str(report_path)]
        environment = dict(os.environ)
        environment.update(TEMP=str(self.work), TMP=str(self.work), PSModulePath=str(self.shell.parent / "Modules"))
        for key in list(environment):
            if key in ("GIT_DIR", "GIT_WORK_TREE", "GIT_INDEX_FILE", "GIT_OBJECT_DIRECTORY", "GIT_ALTERNATE_OBJECT_DIRECTORIES"):
                del environment[key]
        jobs = []

        def factory(command, *, cwd, env):
            job = self.supervisor._default_process_factory(command, cwd=cwd, env=env)
            jobs.append(job)
            return job

        record = {"id": case_id, "executable": str(executable), "arguments": list(map(str, arguments)),
                  "cwd": str(cwd), "deadline_seconds": timeout, "outer_argv": argv,
                  "outer_deadline_seconds": timeout + 30, "actual_process": True}
        self.records.append(record)  # Retain even startup/cancellation/close failures.
        try:
            outer = self.supervisor._run_converter(argv, cwd=str(cwd), env=environment,
                timeout_seconds=timeout + 30, process_factory=factory)
            record["outer_process"] = plain(outer)
        except BaseException as error:
            record.update(outer_error_type=type(error).__name__, outer_error=str(error)[:2048],
                          outer_tree_stopped=getattr(error, "tree_stopped", False),
                          outer_process=plain(getattr(error, "conversion", {})))
            self.timed_out |= getattr(error, "code", "") == "conversion_timeout"
            raise
        record["outer_job"] = {"assigned_before_resume": len(jobs) == 1 and jobs[0].assigned_before_resume,
                               "tree_stopped": len(jobs) == 1 and jobs[0]._tree_stopped,
                               "handles_closed": len(jobs) == 1 and jobs[0]._closed,
                               "pid": jobs[0].process.pid if jobs else None}
        require(all(record["outer_job"][key] is True for key in
                    ("assigned_before_resume", "tree_stopped", "handles_closed")), "Unproved outer owned-job completion")
        require(outer["exit_code"] == 0 and not outer["stdout_truncated"] and not outer["stderr_truncated"]
                and not outer["stderr_tail"], "Outer process probe host failed or produced incomplete output")
        require(report_path.stat().st_size <= RESULT_LIMIT, "Native report exceeds its bounded evidence limit")
        raw = report_path.read_text(encoding="utf-8")
        report = self.json_reader(raw)
        require(type(report) is dict and type(report.get("schema_version")) is int and report["schema_version"] == 1
                and type(report.get("cases")) is list and len(report["cases"]) == 1
                and report["cases"][0].get("id") == case_id and report["cases"][0].get("actual_process") is True
                and report["cases"][0].get("arguments") == payload["cases"][0]["arguments"], "Native case receipt differs")
        native = report["cases"][0]["process"]
        record.update(native_process=native, probe_host={key: report[key] for key in
                      ("host_version", "host_major", "clr", "shell_executable", "stored_policies")},
                      native_receipt_sha256=digest(report_path))
        self.timed_out |= native.get("TimedOut") is True
        validate_process(native)
        return {"exit_code": native["ExitCode"], "stdout": native["Stdout"], "stderr": native["Stderr"],
                "command": [str(executable), *list(map(str, arguments))], "native_case_id": case_id}


INVENTORY_CODE = """import importlib.metadata as m,json,platform,sys,sysconfig
print(json.dumps({'executable':sys.executable,'version':platform.python_version(),'prefix':sys.prefix,
'base_prefix':sys.base_prefix,'purelib':sysconfig.get_path('purelib'),'platlib':sysconfig.get_path('platlib'),
'isolated':bool(sys.flags.isolated),'dont_write_bytecode':bool(sys.dont_write_bytecode),
'distributions':sorted([{'name':d.metadata['Name'],'version':d.version} for d in m.distributions()],key=lambda x:x['name'].casefold())}))
"""
OS_FACTS_CODE = r"""$ErrorActionPreference='Stop'
Import-Module ([IO.Path]::Combine($PSHOME,'Modules','Microsoft.PowerShell.Utility','Microsoft.PowerShell.Utility.psd1')) -ErrorAction Stop
Import-Module ([IO.Path]::Combine($PSHOME,'Modules','CimCmdlets','CimCmdlets.psd1')) -ErrorAction Stop
$os=Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
[ordered]@{caption=$os.Caption;version=$os.Version;build=$os.BuildNumber;architecture=$os.OSArchitecture;
product_type=[int]$os.ProductType;shell_version=$PSVersionTable.PSVersion.ToString();
shell_executable=[Diagnostics.Process]::GetCurrentProcess().MainModule.FileName;clr=[Environment]::Version.ToString()} | ConvertTo-Json -Compress
"""


def single_json(observation, reader):
    require(observation["exit_code"] == 0 and not observation["stderr"], "Native JSON observation failed")
    lines = [line for line in observation["stdout"].splitlines() if line.strip()]
    require(len(lines) == 1, "Exactly one native JSON record is required")
    record = reader(lines[0])
    require(type(record) is dict, "Native JSON observation must be an object")
    return record


def execute_matrix(application_python, shells, work, *, calibre=None):
    evidence = {"schema_version": 1, "task_id": "V101-07", "kind": "native-application-runtime-matrix",
                "observed_at": datetime.now(timezone.utc).isoformat(), "success": False, "timed_out": False,
                "cleanup_safe": False, "driver": {"executable": sys.executable, "version": sys.version,
                    "isolated": bool(sys.flags.isolated), "dont_write_bytecode": bool(sys.dont_write_bytecode)},
                "cases": [], "native_probes": [], "conversion": {"status": "NOT_REQUESTED" if calibre is None else "REQUESTED"},
                "not_asserted": ["native Windows 10", "human Explorer drag/drop", "interactive console acceptance",
                                 "arbitrary ebooks or private documents"]}
    transport = None
    try:
        require(os.name == "nt" and sys.flags.isolated and sys.dont_write_bytecode, "Native Windows isolated driver required")
        candidate = ordinary_file(application_python)
        hosts = [ordinary_file(shell) for shell in shells]
        require(len(hosts) == 2 and len(set(hosts)) == 2, "Both distinct native shell hosts are required")
        if calibre is not None:
            calibre = ordinary_file(calibre)
        manual = load("wbs_matrix_manual", ROOT / "tests/manual/characterize_manual.py")
        launchers = load("wbs_matrix_launchers", ROOT / "tests/manual/current_launchers.py")
        paths = load("wbs_matrix_paths", ROOT / "tests/paths/characterize_paths.py")
        process = load("wbs_matrix_copy", ROOT / "tests/process/characterize_process.py")
        supervisor = load("wbs_matrix_supervisor", ROOT / "engine/winbooksplit_conversion.py")
        compatibility = load("wbs_matrix_policy", ROOT / "engine/winbooksplit_compatibility.py")
        runtime_receipt = load("wbs_matrix_runtime_receipt", ROOT / "tests/runtime/validate_runtime_report.py")
        windows = load("wbs_matrix_windows", ROOT / "engine/winbooksplit_windows.py")
        app = work / "application"
        copied = process.copy_application(app)
        evidence["application_copy"] = {"path": str(app), "source_sha256": copied, "member_count": len(copied)}
        require(len(copied) == 21, "Application fixture must contain the complete 21-file shipped closure")
        policy_path = app / "engine/WinBookSplit.Compatibility.json"
        policy, policy_sha = compatibility.load_policy(policy_path), digest(policy_path)
        evidence["policy"] = {"path": str(policy_path), "sha256": policy_sha, "policy_version": policy["policy_version"]}
        transport = NativeCases(work, app, hosts[0], supervisor, launchers._unique_json)
        evidence["native_probes"] = transport.records
        for index, host in enumerate(hosts):
            arguments = ["-NoProfile", "-NonInteractive", "-EncodedCommand", base64.b64encode(OS_FACTS_CODE.encode("utf-16le")).decode("ascii")]
            facts = single_json(transport.probe(f"host-{index}", host, arguments, work), launchers._unique_json)
            require(type(facts.get("product_type")) is int and facts["product_type"] in (1, 2, 3)
                    and all(type(facts.get(key)) is str and facts[key] for key in
                            ("caption", "version", "build", "architecture", "shell_version", "shell_executable", "clr")),
                    "Actual native OS/shell facts are missing")
            require(Path(facts["shell_executable"]) == host, "Native shell observation differs from the requested literal executable")
            evidence.setdefault("hosts", []).append(facts)
        require({item["shell_version"].split(".")[0] for item in evidence["hosts"]} == {"5", "7"}, "Actual hosts must be PS5.1 and PS7")
        require(all(item["shell_version"].split(".")[:2] == ["5", "1"] for item in evidence["hosts"]
                    if item["shell_version"].split(".")[0] == "5"), "The inbox reference host must actually be PowerShell 5.1")
        inventory = single_json(transport.probe("inventory-before", candidate, ["-I", "-B", "-c", INVENTORY_CODE], work), launchers._unique_json)
        probe = single_json(transport.probe("candidate-full-policy", candidate,
            ["-I", "-B", "-X", "utf8", str(app / "engine/winbooksplit_compatibility.py"), "--probe", "--policy", str(policy_path)], work), launchers._unique_json)
        validate_candidate(probe, inventory, candidate, policy, policy_sha, compatibility)
        evidence["application_runtime"] = {"candidate": str(candidate), "executable_sha256": digest(candidate),
                                           "package_origin_sha256": digest(Path(probe["pypdf_path"])),
                                           "full_probe": probe, "inventory_before": inventory}
        evidence["driver"]["packages"] = {name: metadata.version(name) for name in ("pypdf", "reportlab", "Pillow")}
        units = transport.probe("runtime-only-unit-route", candidate,
            ["-I", "-B", str(ROOT / "tests/run_tests.py"), "--_application-units-child"], work, timeout=300)
        frames = [launchers._unique_json(line[len("[APPLICATION-UNITS] "):]) for line in units["stdout"].splitlines()
                  if line.startswith("[APPLICATION-UNITS] ")]
        require(units["exit_code"] == 0 and len(frames) == 1, "Required runtime-only units failed")
        unit = validate_units(frames[0], candidate, probe["version"])
        evidence["runtime_only_units"] = unit
        generator = manual.history.load_generator()
        fixtures = generator.generate_fixtures(work / "fixtures")
        flat, zero = work / "flat.pdf", work / "zero.pdf"
        flat.write_bytes(generator._pdf_bytes({"title": "Authored no-bookmark fixture", "pages": 3, "outline": []}))
        from pypdf import PdfReader, PdfWriter
        with zero.open("xb") as stream:
            PdfWriter().write(stream)
        sources = [*fixtures.values(), flat, zero]
        source_hashes = {str(path): digest(path) for path in sources}
        evidence["input_sha256_before"] = source_hashes
        definitions = [
            ("manual-front-matter", fixtures["nested12"], "manual", "3,9", [[0, 2], [2, 8], [8, 12]], 12, "split_complete", 0),
            ("auto-level1", fixtures["simple10"], "1", None, [[0, 3], [3, 6], [6, 10]], 10, "split_complete", 0),
            ("auto-level2", fixtures["nested12"], "2", None, [[0, 2], [2, 3], [3, 6], [6, 8], [8, 10], [10, 12]], 12, "split_complete", 0),
            ("late-invalid-token", fixtures["nested12"], "manual", "1,3,bad", None, 12, "invalid_start_pages", 2),
            ("no-bookmarks", flat, "1", None, None, 3, "no_bookmarks", 5),
            ("zero-page", zero, "manual", "1", None, 0, "invalid_document", 6),
        ]
        with ExitStack() as stack:
            for source in sources:
                stack.enter_context(paths.read_only_source(source))
            for case_id, source, mode, starts, ranges, pages, code, exit_code in definitions:
                cwd = work / case_id
                cwd.mkdir()
                output = cwd / "output"
                output.mkdir()
                neighbor = output / "neighbor.txt"
                neighbor.write_bytes(b"Preserve authored neighbor.\n")
                arguments = ["-I", "-B", str(app / "engine/winbooksplit_engine.py"), str(source), str(output), mode]
                if starts is not None:
                    arguments.append(starts)
                observed = transport.probe(case_id, candidate, arguments, cwd)
                frame = manual.result_record(observed)
                require(observed["exit_code"] == exit_code and not observed["stderr"] and frame["code"] == code, "Direct application outcome differs: " + case_id)
                outputs = manual.published_outputs(output, generator, observed)
                if ranges is not None:
                    manual.check_outputs(outputs, ranges, pages)
                    publication = manual.published_directory(output, observed)
                    source_metadata = dict(PdfReader(source).metadata or {})
                    application_version = (app / "VERSION").read_text(encoding="utf-8").strip()
                    for item in outputs:
                        actual_metadata = dict(PdfReader(publication / item["filename"]).metadata or {})
                        require(actual_metadata.get("/Title") == re.sub(r"^[0-9]+ - ", "", Path(item["filename"]).stem)
                                and actual_metadata.get("/Author") == source_metadata.get("/Author")
                                and actual_metadata.get("/Subject") == "Source: " + source_metadata["/Title"]
                                and actual_metadata.get("/Creator") == "WinBookSplit"
                                and actual_metadata.get("/Producer") == f"WinBookSplit {application_version} (pypdf {policy['pypdf']['version']})",
                                "Actual chapter metadata differs from the documented bounded metadata policy")
                        item["metadata"] = actual_metadata
                else:
                    require(not outputs and frame["written_count"] == 0, "Failed application wrote chapter files")
                manual.check_base_members(output, observed, [neighbor.name])
                require(neighbor.read_bytes() == b"Preserve authored neighbor.\n", "Application changed neighbor")
                evidence["cases"].append({"id": case_id, "status": "PASS", "scope": "actual selected application Python",
                    "process_id": observed["native_case_id"], "outcome": frame, "outputs": outputs, "neighbor_unchanged": True})
            for index, host in enumerate(hosts):
                case_id = f"wrapper-PS{evidence['hosts'][index]['shell_version'].split('.')[0]}"
                cwd = work / case_id
                cwd.mkdir()
                output = cwd / "output"
                output.mkdir()
                source = fixtures["nested12"]
                arguments = ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-File", str(app / "WinBookSplit.ps1"),
                    "-InputFile", str(source), "-OutputDirectory", str(output), "-Mode", "Manual", "-StartPages", "3,9",
                    "-PythonPath", str(candidate), "-NonInteractive", "-NoPause"]
                observed = transport.probe(case_id, host, arguments, cwd)
                outcomes = [launchers._unique_json(line[len("[OUTCOME] "):]) for line in observed["stdout"].splitlines() if line.startswith("[OUTCOME] ")]
                require(observed["exit_code"] == 0 and not observed["stderr"] and len(outcomes) == 1
                        and outcomes[0].get("status") == "success" and outcomes[0].get("code") == "split_complete"
                        and outcomes[0].get("exit_code") == 0, "Native wrapper outcome differs")
                outcome = outcomes[0]
                outputs = manual.published_outputs(output, generator, outcome["engine_result"])
                manual.check_outputs(outputs, [[0, 2], [2, 8], [8, 12]], 12)
                logs = [line[len("Log: "):].strip() for line in observed["stdout"].splitlines() if line.startswith("Log: ")]
                require(len(logs) == 1, "Native wrapper omitted its owned local log")
                log = Path(logs[0])
                with windows.FileGuard(log) as log_guard:
                    require(log.stat().st_size <= RESULT_LIMIT, "Wrapper log exceeds the bounded matrix evidence limit")
                    with log.open("rb") as stream:
                        log_bytes = stream.read(RESULT_LIMIT + 1)
                    require(len(log_bytes) <= RESULT_LIMIT, "Wrapper log exceeds the bounded matrix evidence limit")
                    console = launchers.authenticate_console_manifest(log, outcome, source_path=source,
                                                                      expected_python_version=probe["version"])
                    require(len(log_bytes) == console["log_size_bytes"]
                            and hashlib.sha256(log_bytes).hexdigest() == console["log_sha256"],
                            "Wrapper dependency log bytes differ from the authenticated console")
                    dependency = validate_wrapper_runtime(log_bytes.decode("utf-8", errors="strict"), candidate,
                        probe, policy_sha, runtime_receipt, launchers._unique_json)
                    log_guard.assert_unchanged()
                manual.check_base_members(output, outcome["engine_result"], [log.parent.name])
                evidence["cases"].append({"id": case_id, "status": "PASS", "scope": "actual native PowerShell wrapper",
                    "process_id": observed["native_case_id"], "outcome": outcome, "outputs": outputs,
                    "console": console, "dependency": dependency})
        if calibre is not None:
            conversion = load("wbs_matrix_conversion", ROOT / "tests/conversion/characterize_conversion.py")
            fixtures_module = conversion.fixtures
            converter_calls = []

            def owned_conversion(argv, cwd):
                case_id = f"calibre-fixture-{len(converter_calls)}"
                observed = transport.probe(case_id, calibre, list(map(str, argv))[1:], Path(cwd), timeout=180)
                converter_calls.append(observed)
                return {**observed, "argv": list(map(str, argv)), "cwd": str(cwd)}

            with patch.object(fixtures_module, "run", owned_conversion):
                ebook_dir = work / "ebooks"
                provenance = fixtures_module.generate(ebook_dir, calibre)
                ebook_sources = {fmt: ebook_dir / ("original-three-chapters." + fmt) for fmt in ("epub", "azw3")}
                shared_neighbor = conversion.original_neighbor(ebook_sources["epub"], source_hashes)
                require(all(source.with_suffix(".pdf") == shared_neighbor for source in ebook_sources.values()),
                        "Authored ebook neighbor identity differs")
                for source in ebook_sources.values():
                    source_hashes[str(source)] = digest(source)
                neighbor_details = shared_neighbor.lstat()
                neighbor_identity = {"path": str(shared_neighbor), "sha256": source_hashes[str(shared_neighbor)],
                                     "size_bytes": neighbor_details.st_size, "device": neighbor_details.st_dev,
                                     "inode": neighbor_details.st_ino, "attributes": neighbor_details.st_file_attributes}
                ebook_cases = []
                for fmt, keep in (("epub", False), ("azw3", True)):
                    source = ebook_sources[fmt]
                    reference = conversion.reference_pdf(source, work / (fmt + "-reference.pdf"), calibre, work)
                    cwd = work / (fmt + "-application")
                    cwd.mkdir()
                    output = cwd / "output"
                    output.mkdir()
                    arguments = ["-I", "-B", str(app / "engine/winbooksplit_engine.py"), str(source), str(output), "manual", "2,3",
                                 "--calibre-path", str(calibre)]
                    if keep:
                        arguments.append("--keep-converted-pdf")
                    with paths.read_only_source(source):
                        observed = transport.probe(fmt + "-application", candidate, arguments, cwd, timeout=240)
                    frame = manual.result_record(observed)
                    require(observed["exit_code"] == 0 and not observed["stderr"] and frame["code"] == "split_complete", "Actual ebook application failed")
                    execution = frame["execution"]
                    checked, members = conversion.check_publication(output, execution, reference["page_content_sha256"])
                    require((execution["retained_intermediate"] is not None) is keep, "Ebook retention behavior differs")
                    manual.check_base_members(output, observed)
                    case = {"id": fmt + "-application", "status": "PASS", "process_id": observed["native_case_id"],
                            "outcome": frame, "reference": reference, "publication": checked, "owned_members": sorted(members)}
                    ebook_cases.append(case)
                    evidence["cases"].append(case)
                evidence["conversion"] = {"status": "PASS", "scope": "actual locked Calibre EPUB/AZW3 with selected application Python",
                                          "fixture_provenance": provenance, "fixture_calls": converter_calls, "cases": ebook_cases,
                                          "shared_neighbor_identity": neighbor_identity,
                                          "ebook_input_sha256": {fmt: source_hashes[str(source)] for fmt, source in ebook_sources.items()}}
        after = single_json(transport.probe("inventory-after", candidate, ["-I", "-B", "-c", INVENTORY_CODE], work), launchers._unique_json)
        evidence["application_runtime"]["inventory_after"] = after
        require(after == inventory, "Application runtime inventory changed")
        probe_after = single_json(transport.probe("candidate-full-policy-after", candidate,
            ["-I", "-B", "-X", "utf8", str(app / "engine/winbooksplit_compatibility.py"), "--probe", "--policy", str(policy_path)], work), launchers._unique_json)
        validate_candidate(probe_after, after, candidate, policy, policy_sha, compatibility)
        evidence["application_runtime"]["full_probe_after"] = probe_after
        require(probe_after == probe and digest(candidate) == evidence["application_runtime"]["executable_sha256"]
                and digest(Path(probe["pypdf_path"])) == evidence["application_runtime"]["package_origin_sha256"],
                "Selected interpreter/package identity changed")
        input_after = {name: digest(Path(name)) for name in source_hashes}
        evidence["input_sha256_after"] = input_after
        require(input_after == source_hashes, "Authored immutable source or same-name neighbor changed")
        copied_after = {name: digest(app / name) for name in copied}
        evidence["application_copy"]["after_sha256"] = copied_after
        require(copied_after == copied, "Copied application source changed")
        evidence.update(success=True, cleanup_safe=True)
    except BaseException as error:
        evidence.update(error_type=type(error).__name__, error=str(error)[:2048], success=False, cleanup_safe=False)
    finally:
        if transport is not None:
            evidence["timed_out"] = transport.timed_out
    return evidence
