"""Synthetic evidence-validation units; these are not Windows acceptance runs."""

from copy import deepcopy
import importlib.util
import json
import ntpath
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[2]
SPEC = importlib.util.spec_from_file_location("wbs_runtime_receipt_validator",
                                            ROOT / "tests/runtime/validate_runtime_report.py")
validator = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(validator)
SPEC = importlib.util.spec_from_file_location("wbs_runtime_synthetic_session", ROOT / "tests/python/session_receipt_fixture.py")
session_fixture = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(session_fixture)
SPEC = importlib.util.spec_from_file_location("wbs_runtime_console_fixture", ROOT / "tests/python/console_receipt_fixture.py")
console_fixture = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(console_fixture)
SHELLS = [r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe", r"C:\synthetic\PS7\pwsh.exe"]
CALIBRE = r"C:\synthetic\Calibre\ebook-convert.exe"
CONTENT = ["c" * 64, "d" * 64, "e" * 64]
APP_FILES = ("VERSION", "WinBookSplit.ps1", "WinBookSplit.bat", "requirements.txt", "engine/WinBookSplit.Paths.ps1",
             "engine/WinBookSplit.Diagnostics.ps1", "engine/WinBookSplit.Runtime.ps1",
             "engine/WinBookSplit.Process.ps1", "engine/WinBookSplit.Outcomes.json",
             "engine/winbooksplit_engine.py", "engine/winbooksplit_windows.py",
             "engine/winbooksplit_conversion.py", "engine/winbooksplit_job.py", "engine/WinBookSplit.Logging.ps1",
             "engine/WinBookSplit.Support.ps1", "Export-WinBookSplitDiagnostics.ps1",
    "engine/WinBookSplit.Compatibility.ps1", "engine/WinBookSplit.Compatibility.json",
    "engine/winbooksplit_compatibility.py", "Setup-WinBookSplit.ps1", "engine/WinBookSplit.Setup.ps1")


def native_probe(stdout, exit_code=0):
    return {"ExitCode": exit_code, "Pid": 1234, "TimedOut": False, "ParentStopped": True,
            "DescendantsStopped": True, "StreamsComplete": True, "JobAssigned": True, "Cancelled": False,
            "StartError": None, "StopError": None,
            "Stdout": stdout, "Stderr": "", "StdoutTotalBytes": len(stdout.encode()),
            "StderrTotalBytes": 0, "StdoutTruncated": False, "StderrTruncated": False,
            "StreamError": None, "ElapsedSeconds": 0.1}


def selected_runtime(executable, source, version="3.14.8"):
    origin = ntpath.join(ntpath.dirname(ntpath.dirname(executable)),
                         "Lib", "site-packages", "pypdf", "__init__.py")
    details = {"protocol": "winbooksplit.runtime", "schema_version": 1, "ok": True,
               "executable": executable, "version": version, "version_info": [int(part) for part in version.split(".")],
               "releaselevel": "final", "serial": 0, "implementation": "cpython",
               "platform": "win32", "machine": "AMD64", "bits": 64, "gil_disabled": False,
               "isolated": True, "dont_write_bytecode": True, "pypdf_version": "6.20.0",
               "pypdf_path": origin, "policy_version": validator.console.RUNTIME_POLICY["policy_version"],
               "policy_sha256": validator.console.RUNTIME_POLICY_SHA256, "reason": "supported",
               "message": "Synthetic supported interpreter and locked pypdf record."}
    probe = native_probe(json.dumps(details) + "\r\n")
    result = {"Path": executable, "Version": version, "Source": source, "PypdfVersion": "6.20.0",
            "PolicyVersion": details["policy_version"], "PolicySha256": details["policy_sha256"], "InterpreterOnly": False,
            "PypdfPath": origin, "Details": details, "Probe": probe, "Arguments": ["-I", "-B"],
            "Attempts": [{"Path": executable, "Source": source, "Accepted": True,
                          "Reason": "supported", "Message": details["message"], "Probe": deepcopy(probe)}]}
    if source in {"PATH", "manager", "py_launcher"}:
        result["Attempts"].append({"Path": executable, "Source": "selection", "Accepted": True,
                                  "Reason": "supported", "Message": "Synthetic policy selection bookkeeping.", "Probe": None})
    return result


def rejected_details(executable, version, reason):
    details = deepcopy(selected_runtime(executable, "explicit", version)["Details"])
    details.pop("pypdf_version")
    details.pop("pypdf_path")
    details.update(ok=False, reason=reason, message={
        "unsupported_python": "This synthetic Python branch is excluded by the policy.",
        "missing_pypdf": "The locked package cannot be imported: No module named 'pypdf'",
        "pypdf_version_mismatch": "Expected locked pypdf 6.20.0 in both import and distribution metadata.",
    }[reason])
    return details


def selected_converter(source):
    probe = native_probe("ebook-convert.exe (calibre 9.15.0)\r\nCreated by: Kovid Goyal <kovid@kovidgoyal.net>\r\n")
    return {"Path": CALIBRE, "Version": "9.15.0", "Source": source, "Probe": probe,
            "Attempts": [{"Path": CALIBRE, "Source": source, "Accepted": True,
                          "Message": "Supported converter version.", "Probe": deepcopy(probe)}]}


def original_identity(digest="a" * 64):
    return {"sha256": digest, "size_bytes": 200, "device": 1, "inode": 2, "attributes": 32}


def successful_writer(case, ebook):
    source = {"path": case["source_path"], "sha256": "a" * 64, "size_bytes": 200, "binding": "reader_snapshot"}
    entries = [{"sequence": n, "title": f"Section {n}", "start": n - 1, "end": n,
                "parent_id": None, "reason": "manual", "filename": f"{n:02d} - Section {n}.pdf",
                "warnings": [], "page_count": 1, "sha256": f"{n:064x}", "size_bytes": 100 + n}
               for n in range(1, 4)]
    writer = {"schema_version": 1, "status": "complete", "run_id": "1" * 32,
              "final_directory": case["output_base"] + "\\Book_20261009-120000_" + "1" * 32,
              "mode": "manual", "total_pages": 3, "source_identity": source,
              "coverage": {"complete": True, "covered_pages": 3, "section_count": 3},
              "written_count": 3, "outputs": entries, "manifest_filename": "WinBookSplit_Manifest.json"}
    if ebook:
        original = {**source, "binding": "ebook_snapshot"}
        source.update(path=case["output_base"] + "\\.WinBookSplit-stage-" + "2" * 32 + "\\WinBookSplit_Converted.pdf",
                      sha256="f" * 64)
        generated = {**source, "page_count": 3, "page_content_sha256": deepcopy(CONTENT)}
        conversion = {"converter_path": CALIBRE, "original_source_identity": original,
                      "generated_pdf_identity": generated, "exit_code": 0,
                      "workspace_cleanup": {"cleanup_complete": True, "retained_staging": None}}
        writer.update(original_ebook_identity=original, conversion=conversion, retained_intermediate=None)
    writer["manifest"] = {key: deepcopy(value) for key, value in writer.items() if key != "manifest_filename"}
    return writer


def valid_report():
    """An authored complete receipt fixture, with no claim that its processes ran."""
    source_map = {name: f"{n:064x}" for n, name in enumerate(APP_FILES, 1)}
    for name in ("tests/runtime/characterize_runtime.py", "tests/runtime/validate_runtime_report.py",
                 "tests/runtime/converter_cache.py", "tests/runtime/Observe-ManagerCache.ps1",
                 "tests/runtime/README.md", "tests/python/test_runtime_receipt.py",
                 "tests/baseline/expected_original.json", "tests/extraction/characterize_extraction.py",
                 "tests/fixtures/generate_pdf_fixtures.py"):
        source_map[name] = "b" * 64
    policies = [{"scope": name, "policy": "Undefined"} for name in
                ("MachinePolicy", "UserPolicy", "CurrentUser", "LocalMachine")]
    hosts = []
    for name, shell, major, version in zip(("PS51", "PS7"), SHELLS, (5, 7), ("5.1.26100.9444", "7.6.6")):
        observed = {"host_major": major, "host_version": version, "syntax_error_count": 0,
                    "syntax_checked": ["WinBookSplit.ps1", "engine/WinBookSplit.Paths.ps1",
                                       "engine/WinBookSplit.Diagnostics.ps1", "engine/WinBookSplit.Runtime.ps1", "engine/WinBookSplit.Process.ps1"],
                    "stored_policies": deepcopy(policies)}
        hosts.append({"id": name, "passed": True, "shell_executable": shell, "exit_code": 0,
                      "command": [shell, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned",
                                  "-File", r"C:\synthetic\Host.ps1"], "cwd": r"C:\synthetic",
                      "stdout": json.dumps(observed), "stderr": "", **observed, "policies_after": deepcopy(policies)})
    report = {"schema_version": 1, "task_id": "M2-T04", "result": "RUNTIME_REGRESSION_PASSED",
              "success": True, "exit_code": 0, "acceptance_ids": ["AC-043", "AC-044", "AC-045"],
              "host_cases": hosts, "source_unchanged": True, "baseline_guards_preserved": True,
              "input_and_neighbor_unchanged": True, "machine_settings_unchanged": True, "owned_temp_removed": True,
              "immutable_original_commit": "0de84f367f9bd5ddfa3f408a9c29505d7a39633f",
              "tested_path_sha256": source_map, "case_count": 33,
              "calibre": {"path": CALIBRE, "version": "9.15.0",
                          "sha256": "f46a01c9b8cd392e1190ce2e50a2a5a48b326c8c3eab670bbcd8b298bede501d"}}
    groups = {
        "python_failure_cases": ("python_failure", ("missing", "non-python", "wrong-version", "missing-pypdf", "wrong-pypdf")),
        "python_selection_cases": ("python_selection", ("explicit-literal", "app-venv", "multiple-path", "py-launcher-listing")),
        "shadow_cases": ("shadow", ("book-cwd-env-decoys",)),
        "pdf_without_calibre_cases": ("pdf_without_calibre", ("pdf-no-calibre",)),
        "converter_failure_cases": ("converter_failure", ("missing", "wrong-version", "book-cwd-decoys")),
        "converter_success_cases": ("converter_success", ("explicit-portable", "path-portable")),
    }
    number = 0
    for field, (category, kinds) in groups.items():
        report[field] = []
        pairs = [(host, kind) for host in ("PS51", "PS7") for kind in kinds]
        if field == "converter_success_cases":
            pairs.append(("BAT", "path-portable"))
        for host, kind in pairs:
            number += 1
            directory = rf"C:\synthetic\c{number}"
            ebook = category in {"python_failure", "converter_failure", "converter_success"}
            source = directory + ("\\d\\Book.epub" if ebook else "\\d\\Book.pdf")
            neighbor = directory + ("\\d\\Book.pdf" if ebook else "\\d\\neighbor.txt")
            shell = SHELLS[0 if host in {"PS51", "BAT"} else 1]
            command = ([r"C:\Windows\System32\cmd.exe", "/d", "/v:off", "/c", directory + "\\Invoke.cmd"] if host == "BAT"
                       else [shell, "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", directory + "\\Invoke.ps1"])
            prefix = "python-" if category == "python_failure" else "converter-" if category == "converter_failure" else ""
            originals = {source: original_identity(), neighbor: original_identity("b" * 64)}
            case = {"id": host + "-" + prefix + kind, "category": category, "kind": kind, "passed": True,
                    "actual_process": True, "shell_executable": shell, "command": command, "cwd": directory + "\\c",
                    "output_base": directory + "\\o", "source_path": source, "source_format": "epub" if ebook else "pdf",
                    "source_read_only_attribute_observed": True, "source_attributes_restored": True,
                    "input_and_neighbor_identity": originals, "input_and_neighbor_identity_after": deepcopy(originals),
                    "input_unchanged": True, "neighbor_unchanged": True, "owned_outputs_removed": True,
                    "application_path_sha256": {name: source_map[name] for name in APP_FILES},
                    "copied_application_unchanged": True, "decoy_sha256": {}, "decoy_sha256_after": {},
                    "decoys_unchanged": True, "decoy_execution_marker_absent": True, "module_import_marker_absent": True,
                    "decoy_marker_observation": {f"{name}_{suffix}": (directory + "\\" + name + ".txt" if suffix == "path" else True)
                                                for name in ("native", "module", "converter_version", "conversion")
                                                for suffix in ("path", "absent")},
                    "controlled_PATH": r"C:\synthetic\env\Scripts;C:\Windows\System32",
                    "controlled_PYTHONPATH": None, "stdout": "", "stderr": ""}
            if category == "shadow":
                case["controlled_PYTHONPATH"] = directory + "\\e"
                for folder in ("d", "c", "e"):
                    case["decoy_sha256"][directory + "\\" + folder + "\\pypdf.py"] = "c" * 64
                for folder in ("d", "c"):
                    for filename in ("python.exe", "py.exe"):
                        case["decoy_sha256"][directory + "\\" + folder + "\\" + filename] = "d" * 64
            if category == "converter_failure" and kind == "book-cwd-decoys":
                for folder in ("d", "c"):
                    case["decoy_sha256"][directory + "\\" + folder + "\\ebook-convert.exe"] = "d" * 64
            if category == "converter_failure" and kind in {"missing", "book-cwd-decoys"}:
                identities = {name: {"device": 1, "inode": index} for index, name in enumerate(
                    ("Microsoft", "Microsoft/Windows", "Microsoft/Windows/Caches"), 1)}
                case["converter_absence_environment"] = {
                    "schema_version": 2,
                    "child_local_app_data_isolated": True, "windows_cache_inventory_unchanged": True,
                    "ordinary_directory_guards_unchanged": True,
                    "cache_directory_identities_before": identities,
                    "cache_directory_identities_after": deepcopy(identities),
                    "parent_local_app_data_unchanged": True, "fixed_system_calibre_paths_absent": True,
                    "manager_cache_observation": {"before": None, "after": None, "reference": None}}
            case["decoy_sha256_after"] = deepcopy(case["decoy_sha256"])
            if category in {"python_failure", "converter_failure"}:
                explicit = category == "python_failure" or kind == "wrong-version"
                candidate = directory + "\\candidate.exe" if explicit else None
                attempts = []
                if explicit:
                    detail = rejected_details(candidate, "3.12.14" if kind == "wrong-version" else "3.14.8",
                              "unsupported_python" if kind == "wrong-version" else
                              "missing_pypdf" if kind == "missing-pypdf" else "pypdf_version_mismatch")
                    bad_probe = None if kind == "missing" else native_probe(
                        "ebook-convert.exe (calibre 9.14.0)\r\n" if category == "converter_failure"
                        else "OWNED-NON-PYTHON-NATIVE-CONTROL\r\n" if kind == "non-python" else json.dumps(detail) + "\r\n",
                        0 if category == "converter_failure" else 7 if kind == "non-python" else 1)
                    attempts = [{"Path": candidate, "Source": "explicit", "Accepted": False,
                                 "Message": "Rejected authored control", "Probe": bad_probe}]
                elif kind == "book-cwd-decoys":
                    attempts = [{"Path": filename, "Source": "PATH", "Accepted": False,
                                 "Message": "Excluded automatic candidate", "Probe": None} for filename in case["decoy_sha256"]]
                error = {"Code": "runtime_invalid" if category == "python_failure" else
                         "converter_invalid" if kind == "wrong-version" else "converter_not_found", "Attempts": attempts}
                guidance = ("Reviewed Python policy 6.20.0 -PythonPath Setup-WinBookSplit.ps1 .venv " + candidate if category == "python_failure"
                            else "Calibre 9.15.0 -CalibrePath")
                case.update(dependency_error=error, candidate_path=candidate, exit_code=3,
                            stdout="Dependency preflight failed " + guidance + "\n[DEPENDENCY-ERROR] " + json.dumps(error) + "\n",
                            preflight_before_output=True, output_base_before=[], output_base_after=[],
                            conversion_started=False, converter_version_marker_absent=not (category == "converter_failure" and kind == "wrong-version"),
                            no_success_summary=True, written_count=0, outputs=[], setup_guidance_verified=True,
                            engine_invocation_record_absent=True, dependency_success_record_absent=True)
                case["decoy_marker_observation"]["converter_version_absent"] = case["converter_version_marker_absent"]
            else:
                selection = {"explicit-literal": "explicit", "app-venv": "application_venv",
                             "py-launcher-listing": "py_launcher"}.get(kind, "PATH")
                executable = directory + "\\a\\.venv\\Scripts\\python.exe" if kind == "app-venv" else r"C:\synthetic\env\Scripts\python.exe"
                runtime = selected_runtime(executable, selection)
                if category == "shadow":
                    runtime["Attempts"][:0] = [{"Path": name, "Source": "PATH", "Accepted": False,
                        "Message": "Excluded book/CWD executable", "Probe": None}
                        for name in case["decoy_sha256"] if name.endswith(".exe")]
                if kind == "multiple-path":
                    runtime["Attempts"].insert(0, {"Path": r"C:\synthetic\unsupported\python.exe", "Source": "PATH",
                        "Accepted": False, "Message": "Excluded Python", "Probe": native_probe(json.dumps(
                            rejected_details(r"C:\synthetic\unsupported\python.exe", "3.12.14", "unsupported_python")) + '\r\n', 1)})
                if kind == "py-launcher-listing":
                    launcher = directory + "\\launcher\\py.exe"
                    runtime["Attempts"].insert(0, {"Path": launcher, "Source": "py_launcher_listing", "Accepted": True,
                        "Message": "Read-only listing", "Probe": native_probe(" -V:3.14 * " + executable + "\r\n")})
                    case["launcher_listing_control"] = {"scope": "authored-native-py-listing-control; actual selected interpreter",
                        "path": launcher, "sha256": "a" * 64, "arguments": ["-0p"], "listed_python_path": executable,
                        "marker_path": directory + "\\listing.txt", "marker_sha256": "b" * 64, "control_reached": True}
                converter_source = "explicit" if kind == "explicit-portable" else "PATH" if ebook else None
                writer = successful_writer(case, ebook)
                engine_argv = ["-I", "-B", "-X", "utf8", directory + "\\a\\engine\\winbooksplit_engine.py",
                               source, case["output_base"], "manual", "", "--interactive", "f" * 32]
                if ebook:
                    engine_argv += ["--calibre-path", CALIBRE, "--conversion-timeout", "1800"]
                case.update(dependency_selection={"Runtime": runtime,
                            "Converter": selected_converter(converter_source) if ebook else None},
                            expected_python_path=executable, expected_python_source=selection,
                            expected_python_version="3.14.8",
                            expected_converter_source=converter_source,
                            engine_invocation={"path": executable, "arguments": engine_argv},
                            writer_result=writer, exit_code=0, written_count=3, selected_interpreter_used_for_engine=True,
                            all_physical_pages_preserved=True, manifest_validated=True,
                            independent_source_page_content_sha256=deepcopy(CONTENT), page_content_sha256=deepcopy(CONTENT),
                            outputs=[{"filename": e["filename"], "range": [e["start"], e["end"]],
                                      "page_ids": [e["sequence"]], "sha256": e["sha256"]} for e in writer["outputs"]])
                case["engine_record"] = {"protocol": "winbooksplit.result", "version": 1, "mode": "manual",
                    "status": "success", "code": "split_complete", "exit_code": 0, "written_count": 3,
                    "execution": deepcopy(writer)}
                case["interaction"] = session_fixture.make(writer, case["engine_invocation"], ["input_ready", "plan_ready"], ["starts", "execute"])
                case["process_summaries"] = [{"InteractionError": None, "InputError": None, "InputWriterStopped": True,
                    "QueuedReplyCount": 2, "InteractionCount": 2, "ReplyCount": 2}]
                case["stdin_utf8"] = "M\nN\n2,3\nY\nN\n\n" if ebook else "M\n2,3\nY\nN\n\n"
                final = {"protocol": "winbooksplit.outcome", "version": 1, "mode": "manual", "status": "success",
                    "code": "split_complete", "exit_code": 0, "written_count": 3, "final_directory": writer["final_directory"], "engine_result": case["engine_record"]}
                case["console_evidence"] = console_fixture.make(final, "c" * 64,
                    plan=case["interaction"]["displayed_plans"][-1]["plan"])
            report[field].append(case)
    report["positive_decoy_controls"] = {
        "native": {"command": [r"C:\synthetic\Stub.exe", "--version"], "cwd": r"C:\synthetic", "stdout": "", "stderr": "",
                   "exit_code": 0, "marker_path": r"C:\synthetic\native.txt", "marker_sha256": "a" * 64, "control_reached": True},
        "module": {"command": [r"C:\synthetic\core\python.exe", "-B", "-c", "import pypdf"], "cwd": r"C:\synthetic",
                   "stdout": "", "stderr": "Rejected marker control", "exit_code": 1, "marker_path": r"C:\synthetic\module.txt",
                   "marker_sha256": "b" * 64, "module_path": r"C:\synthetic\module\pypdf.py", "module_sha256": "c" * 64,
                   "controlled_PYTHONPATH": r"C:\synthetic\module", "control_reached": True}}
    report["native_control_compilation"] = {"command": [SHELLS[0], "-Command", "Synthetic fixture"],
        "cwd": r"C:\synthetic", "exit_code": 0, "stdout": "", "stderr": "", "source_sha256": "c" * 64, "executable_sha256": "d" * 64}
    source_map["engine/WinBookSplit.Compatibility.json"] = validator.console.RUNTIME_POLICY_SHA256
    for field in validator.CASE_KINDS:
        for case in report[field]:
            case["application_path_sha256"]["engine/WinBookSplit.Compatibility.json"] = validator.console.RUNTIME_POLICY_SHA256
    helper, policy = str(ROOT / 'engine/winbooksplit_compatibility.py'), str(ROOT / 'engine/WinBookSplit.Compatibility.json')
    supported = selected_runtime(r"C:\synthetic\env\Scripts\python.exe", "explicit")
    report["supported_real_python_observation"] = {"exit_code": 0, "stdout": json.dumps(supported['Details']), "stderr": "",
        "command": [supported['Path'], '-I', '-B', helper, '--probe', '--policy', policy]}
    excluded = rejected_details(r"C:\synthetic\unsupported\python.exe", "3.12.14", "unsupported_python")
    report["unsupported_real_python_observation"] = {"exit_code": 1, "stdout": json.dumps(excluded), "stderr": "",
        "command": [excluded['executable'], '-I', '-B', helper, '--probe', '--interpreter-only', '--policy', policy]}
    report["independent_real_conversion_reference"] = {"process": {"exit_code": 0, "argv": [CALIBRE,
        r"C:\synthetic\reference.epub", r"C:\synthetic\reference.pdf", "--output-profile", "tablet"]},
        "sha256": "f" * 64, "size_bytes": 300, "page_count": 3, "page_content_sha256": deepcopy(CONTENT)}
    return report


class RuntimeReceiptTests(unittest.TestCase):
    def validate(self, report):
        validator.validate_runtime_report(report, SHELLS, CALIBRE)

    def test_complete_synthetic_receipt_is_structurally_accepted(self):
        self.validate(valid_report())

    def test_converter_absence_cache_proof_rejects_omissions_and_corruption(self):
        def absence(report): return report["converter_failure_cases"][0]["converter_absence_environment"]
        mutations = {
            "missing proof": lambda r: r["converter_failure_cases"][0].pop("converter_absence_environment"),
            "missing decoy proof": lambda r: r["converter_failure_cases"][2].pop("converter_absence_environment"),
            "false parent preservation": lambda r: absence(r).update(parent_local_app_data_unchanged=False),
            "numeric parent promise": lambda r: absence(r).update(parent_local_app_data_unchanged=1),
            "unheld layout": lambda r: absence(r).update(ordinary_directory_guards_unchanged=False),
            "system converter present": lambda r: absence(r).update(fixed_system_calibre_paths_absent=False),
            "manager observation omitted": lambda r: absence(r).pop("manager_cache_observation"),
            "legacy whole-cache empty claim": lambda r: absence(r).update(owned_empty_cache_inventory_unchanged=True),
            "root claimed empty": lambda r: absence(r).update(owned_directory_ordinary_and_empty=True),
            "Programs install": lambda r: absence(r)["cache_directory_identities_before"].update(Programs={"device": 1, "inode": 8}),
            "Calibre sibling": lambda r: absence(r)["cache_directory_identities_after"].update({"Microsoft/Windows/Calibre": {"device": 1, "inode": 8}}),
            "cache file": lambda r: absence(r)["cache_directory_identities_after"].update({"Microsoft/Windows/Caches/unknown.dat": {"device": 1, "inode": 8}}),
            "missing cache leaf": lambda r: absence(r)["cache_directory_identities_before"].pop("Microsoft/Windows/Caches"),
            "replaced inode": lambda r: absence(r)["cache_directory_identities_after"]["Microsoft"].update(inode=8),
            "bool device": lambda r: absence(r)["cache_directory_identities_before"]["Microsoft"].update(device=True),
            "invalid inode": lambda r: absence(r)["cache_directory_identities_before"]["Microsoft"].update(inode=-1),
            "unexpected identity field": lambda r: absence(r)["cache_directory_identities_before"]["Microsoft"].update(path="../Programs"),
        }
        for label, mutate in mutations.items():
            with self.subTest(defect=label):
                report = valid_report()
                mutate(report)
                # Keep forged layouts internally equal so closure/type guards
                # are exercised independently of before/after equality.
                proof = absence(report) if "converter_absence_environment" in report["converter_failure_cases"][0] else None
                if label in {"Programs install", "bool device", "invalid inode", "unexpected identity field"}:
                    proof["cache_directory_identities_after"] = deepcopy(proof["cache_directory_identities_before"])
                elif label in {"Calibre sibling", "cache file"}:
                    proof["cache_directory_identities_before"] = deepcopy(proof["cache_directory_identities_after"])
                with self.assertRaises(ValueError): self.validate(report)

    def test_actual_authored_cache_inventory_rejects_unexpected_entries(self):
        spec = importlib.util.spec_from_file_location("wbs_runtime_cache_fixture", ROOT / "tests/runtime/characterize_runtime.py")
        producer = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(producer)
        for unexpected in (None, "Programs/Calibre", "Microsoft/Windows/Calibre", "Microsoft/Windows/Caches/unknown.dat", "Other"):
            with self.subTest(unexpected=unexpected), tempfile.TemporaryDirectory(prefix="WBS-owned-cache-") as owned:
                root = Path(owned).resolve(strict=True) / "cache"
                self.assertTrue(root.is_absolute() and root.is_relative_to(Path(owned).resolve(strict=True)))
                (root / "Microsoft/Windows/Caches").mkdir(parents=True)
                before = producer.converter_absence_cache_inventory(root)
                self.assertEqual(set(before["windows_directories"]), {"Microsoft", "Microsoft/Windows", "Microsoft/Windows/Caches"})
                if unexpected is not None:
                    target = root / unexpected
                    self.assertTrue(target.is_relative_to(root))
                    if target.suffix:
                        target.write_bytes(b"Authored unexpected cache file")
                    else:
                        target.mkdir(parents=True)
                    with self.assertRaises(RuntimeError): producer.converter_absence_cache_inventory(root)
                else:
                    self.assertEqual(producer.converter_absence_cache_inventory(root), before)

    def test_synthetic_policy_versions_and_manager_selection_are_supported(self):
        for version in ('3.13.16', '3.13.17', '3.13.99', '3.14.8', '3.14.10', '3.14.99'):
            for source in ('explicit', 'PATH', 'manager', 'py_launcher'):
                with self.subTest(version=version, source=source):
                    validator.runtime(selected_runtime(r'C:\synthetic\env\Scripts\python.exe', source, version))

    def test_synthetic_previous_security_floors_reject_for_every_discovery_source(self):
        for version in ('3.13.15', '3.14.7'):
            for source in ('explicit', 'PATH', 'manager', 'py_launcher'):
                with self.subTest(version=version, source=source), self.assertRaises(ValueError):
                    validator.runtime(selected_runtime(r'C:\synthetic\env\Scripts\python.exe', source, version))

    def test_synthetic_blocked_patch_and_manifest_target_mismatch_reject(self):
        policy = deepcopy(validator.console.RUNTIME_POLICY)
        next(branch for branch in policy['python']['branches'] if branch['minor'] == 14)['blocked_patches'] = [8]
        with patch.object(validator.console, 'RUNTIME_POLICY', policy), self.assertRaises(ValueError):
            validator.runtime(selected_runtime(r'C:\synthetic\env\Scripts\python.exe', 'PATH'))
        case = valid_report()['python_selection_cases'][0]
        receipt = case['console_evidence']
        for version in ('3.13.16', '3.14.10'):
            receipt['run_manifest']['runtime_versions']['python'] = version
            console_fixture.reseal(receipt)
            validator.console.validate_console_evidence(receipt, receipt['operation_outcome'], receipt['log_sha256'])
            with self.assertRaises(ValueError):
                validator.console.validate_console_evidence(receipt, receipt['operation_outcome'], receipt['log_sha256'],
                                                          expected_python_version='3.14.8')

    def test_resealed_synthetic_identity_policy_and_candidate_corruption_reject(self):
        def reseal(record):
            record['Probe'] = native_probe(json.dumps(record['Details']) + '\r\n')
            record['Attempts'][0]['Probe'] = deepcopy(record['Probe'])
        mutations = {
            'unreviewed branch': lambda r: r['Details'].update(version='3.15.0', version_info=[3, 15, 0]),
            'security floor': lambda r: r['Details'].update(version='3.14.6', version_info=[3, 14, 6]),
            'boolean version': lambda r: r['Details'].update(version_info=[3, 14, True]),
            'prerelease': lambda r: r['Details'].update(releaselevel='candidate', serial=1),
            'policy identity': lambda r: r['Details'].update(policy_sha256='a' * 64),
            'policy boolean': lambda r: r['Details'].update(policy_version=True),
            'unknown native field': lambda r: r['Details'].update(unrecognized=True),
            'unknown provenance': lambda r: r.update(Source='selection'),
            'wrong selected candidate': lambda r: r['Attempts'][0].update(Path=r'C:\other\python.exe'),
            'bookkeeping probe': lambda r: r['Attempts'][-1].update(Probe=deepcopy(r['Probe'])),
            'bookkeeping path': lambda r: r['Attempts'][-1].update(Path=r'C:\other\python.exe'),
            'interpreter only': lambda r: r.update(InterpreterOnly=True),
        }
        for name, mutate in mutations.items():
            with self.subTest(defect=name):
                record = selected_runtime(r'C:\synthetic\env\Scripts\python.exe', 'manager')
                mutate(record)
                reseal(record)
                with self.assertRaises(ValueError):
                    validator.runtime(record)

    def test_windows_host_casing_preserves_identity(self):
        report = valid_report()
        report["host_cases"][0]["shell_executable"] = SHELLS[0].upper()
        self.validate(report)

    def test_missing_and_mislabeled_required_cases_are_rejected(self):
        for field in ("python_failure_cases", "python_selection_cases", "shadow_cases", "pdf_without_calibre_cases",
                      "converter_failure_cases", "converter_success_cases"):
            with self.subTest(field=field):
                report = valid_report()
                report[field].pop()
                with self.assertRaises(ValueError):
                    self.validate(report)
                report = valid_report()
                report[field][0]["id"] = report[field][1]["id"]
                with self.assertRaises(ValueError):
                    self.validate(report)

    def test_promised_evidence_omissions_and_contradictions_are_rejected(self):
        def success(r): return r["python_selection_cases"][0]
        def failure(r): return r["python_failure_cases"][0]
        def shadow(r): return r["shadow_cases"][0]
        def converter_case(r): return r["converter_success_cases"][0]
        mutations = {
            "source guard": lambda r: r.pop("source_unchanged"),
            "historical identity": lambda r: r.update(immutable_original_commit="0" * 40),
            "source map": lambda r: r["tested_path_sha256"].pop("engine/WinBookSplit.Runtime.ps1"),
            "count only": lambda r: r.update(case_count=32),
            "fake native status": lambda r: failure(r).update(exit_code=0),
            "preflight stage": lambda r: failure(r).update(output_base_after=["unknown"]),
            "conversion started": lambda r: failure(r).update(conversion_started=True),
            "missing error frame": lambda r: failure(r).update(stdout="Dependency preflight failed"),
            "mislabeled missing dependency": lambda r: r["python_failure_cases"][3]["dependency_error"]["Attempts"][0]["Probe"].update(
                Stdout='{"protocol":"winbooksplit.runtime","schema_version":1,"ok":false,"version":"3.14.7","message":"Expected CPython 3.14.8 exactly."}'),
            "hidden fallback": lambda r: failure(r)["dependency_error"]["Attempts"].append(deepcopy(failure(r)["dependency_error"]["Attempts"][0])),
            "selected runtime": lambda r: success(r).pop("dependency_selection"),
            "wrong Python": lambda r: success(r)["dependency_selection"]["Runtime"].update(Version="3.14.7"),
            "wrong pypdf": lambda r: success(r)["dependency_selection"]["Runtime"].update(PypdfVersion="9.99.0"),
            "singleton probe": lambda r: success(r)["dependency_selection"]["Runtime"]["Details"].update(protocol=["winbooksplit.runtime"]),
            "boolean schema": lambda r: success(r)["dependency_selection"]["Runtime"]["Details"].update(schema_version=True),
            "native probe mismatch": lambda r: success(r)["dependency_selection"]["Runtime"]["Probe"].update(Stdout="{}"),
            "probe stream absent": lambda r: success(r)["dependency_selection"]["Runtime"]["Probe"].pop("StreamsComplete"),
            "unproved descendants": lambda r: success(r)["dependency_selection"]["Runtime"]["Probe"].update(DescendantsStopped=None),
            "engine changed interpreter": lambda r: success(r)["engine_invocation"].update(path=r"C:\other\python.exe"),
            "engine isolation missing": lambda r: success(r)["engine_invocation"]["arguments"].__setitem__(0, "-B"),
            "unrequested host": lambda r: r["host_cases"][1].update(shell_executable=r"C:\other\pwsh.exe"),
            "policy mutation": lambda r: r["host_cases"][0]["policies_after"][0].update(policy="Bypass"),
            "source replaced": lambda r: success(r)["input_and_neighbor_identity_after"][success(r)["source_path"]].update(inode=9),
            "neighbor missing": lambda r: success(r)["input_and_neighbor_identity"].pop(next(k for k in success(r)["input_and_neighbor_identity"] if k != success(r)["source_path"])),
            "readonly unobserved": lambda r: success(r).update(source_read_only_attribute_observed=False),
            "attributes unrestored": lambda r: success(r).update(source_attributes_restored=False),
            "decoy count only": lambda r: shadow(r).update(decoy_sha256={}, decoy_sha256_after={}),
            "executed decoy": lambda r: shadow(r)["decoy_marker_observation"].update(native_absent=False),
            "changed decoy": lambda r: shadow(r)["decoy_sha256_after"].update({next(iter(shadow(r)["decoy_sha256_after"])): "f" * 64}),
            "unproved positive control": lambda r: r["positive_decoy_controls"]["module"].update(control_reached=False),
            "source content absent": lambda r: success(r).pop("independent_source_page_content_sha256"),
            "content mismatch": lambda r: success(r)["page_content_sha256"].__setitem__(1, "a" * 64),
            "page duplicated": lambda r: success(r)["outputs"][1].update(page_ids=[1]),
            "page boolean": lambda r: success(r)["outputs"][0].update(page_ids=[True]),
            "writer absent": lambda r: success(r).pop("writer_result"),
            "manifest absent": lambda r: success(r)["writer_result"].pop("manifest"),
            "manifest count mismatch": lambda r: success(r)["writer_result"]["manifest"].update(written_count=0),
            "manifest boolean sequence": lambda r: success(r)["writer_result"]["manifest"]["outputs"][0].update(sequence=True),
            "unknown cleanup": lambda r: success(r).update(owned_outputs_removed=False),
            "PDF probes Calibre": lambda r: r["pdf_without_calibre_cases"][0]["dependency_selection"].update(Converter=selected_converter("explicit")),
            "wrong converter": lambda r: converter_case(r)["dependency_selection"]["Converter"].update(Version="9.14.0"),
            "converter trailing spoof": lambda r: converter_case(r)["dependency_selection"]["Converter"]["Probe"].update(Stdout="ebook-convert.exe (calibre 9.15.0)\nEVIL"),
            "ebook snapshot lost": lambda r: converter_case(r)["writer_result"].pop("original_ebook_identity"),
            "conversion workspace retained": lambda r: converter_case(r)["writer_result"]["conversion"]["workspace_cleanup"].update(retained_staging="unknown"),
            "reference absent": lambda r: r.pop("independent_real_conversion_reference"),
            "launcher control absent": lambda r: r["python_selection_cases"][3].pop("launcher_listing_control"),
            "BAT mislabeled": lambda r: r["converter_success_cases"][-1]["command"].__setitem__(1, "/v:on"),
        }
        for name, mutate in mutations.items():
            with self.subTest(defect=name):
                report = valid_report()
                mutate(report)
                with self.assertRaises(ValueError):
                    self.validate(report)


if __name__ == "__main__":
    unittest.main()
