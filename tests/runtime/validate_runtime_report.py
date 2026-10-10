"""Require complete runtime acceptance evidence; this does not authenticate a receipt."""

from __future__ import annotations

import json
import math
import ntpath
from pathlib import PureWindowsPath
import re
import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location("wbs_runtime_interaction_validation", Path(__file__).resolve().parents[2] / "tests/manual/interaction_receipts.py")
interactions = importlib.util.module_from_spec(spec)
spec.loader.exec_module(interactions)
spec = importlib.util.spec_from_file_location("wbs_runtime_console_validation", Path(__file__).resolve().parents[2] / "tests/manual/current_launchers.py")
console = importlib.util.module_from_spec(spec)
spec.loader.exec_module(console)


HOSTS = ("PS51", "PS7")
ACCEPTANCE = ["AC-043", "AC-044", "AC-045"]
ORIGINAL_COMMIT = "0de84f367f9bd5ddfa3f408a9c29505d7a39633f"
CASE_KINDS = {
    "python_failure_cases": ("missing", "non-python", "wrong-version", "missing-pypdf", "wrong-pypdf"),
    "python_selection_cases": ("explicit-literal", "app-venv", "multiple-path", "py-launcher-listing"),
    "shadow_cases": ("book-cwd-env-decoys",),
    "pdf_without_calibre_cases": ("pdf-no-calibre",),
    "converter_failure_cases": ("missing", "wrong-version", "book-cwd-decoys"),
    "converter_success_cases": ("explicit-portable", "path-portable"),
}


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def mapping(value, label: str) -> dict:
    require(type(value) is dict, f"{label}: expected an object")
    return value


def integer(value, label: str, minimum: int = 0) -> int:
    require(type(value) is int and value >= minimum, f"{label}: invalid integer")
    return value


def text(value, label: str) -> str:
    require(type(value) is str and bool(value), f"{label}: missing text")
    return value


def flag(record: dict, field: str, expected: bool = True) -> None:
    require(record.get(field) is expected, f"{field}: missing or contradictory evidence")


def digest(value, label: str) -> str:
    require(type(value) is str and re.fullmatch(r"[0-9a-f]{64}", value) is not None,
            f"{label}: invalid SHA256")
    return value


def path(value, label: str) -> str:
    value = text(value, label)
    require(PureWindowsPath(value).is_absolute(), f"{label}: expected an absolute Windows path")
    return ntpath.normcase(ntpath.normpath(value))


def same_path(first, second, label: str) -> None:
    require(path(first, label) == path(second, label), f"{label}: paths differ")


def same_json(first, second) -> bool:
    return json.dumps(first, sort_keys=True, separators=(",", ":")) == json.dumps(second, sort_keys=True, separators=(",", ":"))


def case_set(child: dict, field: str) -> dict[str, dict]:
    cases = child.get(field)
    require(type(cases) is list, f"{field}: missing cases")
    prefix = "python-" if field == "python_failure_cases" else "converter-" if field == "converter_failure_cases" else ""
    expected = {host + "-" + prefix + kind for host in HOSTS for kind in CASE_KINDS[field]}
    if field == "converter_success_cases":
        expected.add("BAT-path-portable")
    identifiers = [mapping(case, field).get("id") for case in cases]
    require(all(type(identifier) is str for identifier in identifiers), f"{field}: invalid IDs")
    require(len(identifiers) == len(expected) and set(identifiers) == expected,
            f"{field}: exact cases required, without duplicates")
    for case in cases:
        flag(case, "passed")
        flag(case, "actual_process")
        category = field.removesuffix("_cases")
        require(case.get("category") == category and case.get("kind") in CASE_KINDS[field],
                f"{field}: case category/kind mislabeled")
        require(case["id"] == ("BAT" if case["id"].startswith("BAT-") else case["id"].split("-")[0]) +
                "-" + prefix + case["kind"], f"{field}: ID and kind differ")
    return {case["id"]: case for case in cases}


def probe(record, *, success: bool) -> dict:
    record = mapping(record, "probe")
    require(record.get("DescendantsStopped") is True,
            "probe: descendant stop is unproved")
    for field in ("TimedOut", "ParentStopped", "StreamsComplete", "StdoutTruncated", "StderrTruncated"):
        require(type(record.get(field)) is bool, f"probe.{field}: missing boolean")
    for stream in ("Stdout", "Stderr"):
        require(type(record.get(stream)) is str, f"probe.{stream}: missing stream")
        count = integer(record.get(stream + "TotalBytes"), "probe." + stream + "TotalBytes")
        require(record[stream + "Truncated"] is (count > 65536), f"probe.{stream}: truncation mismatch")
        encoded_size = len(record[stream].encode("utf-8"))
        require(encoded_size <= 65542, f"probe.{stream}: unbounded tail")
        if not record[stream + "Truncated"]:
            require(encoded_size == count, f"probe.{stream}: byte count mismatch")
    elapsed = record.get("ElapsedSeconds")
    require(type(elapsed) in (int, float) and math.isfinite(elapsed) and 0 <= elapsed <= 61,
            "probe: unbounded or invalid elapsed time")
    if success:
        integer(record.get("Pid"), "probe.Pid", 1)
        require(type(record.get("ExitCode")) is int and record["ExitCode"] == 0,
                "probe: successful native status required")
        flag(record, "TimedOut", False)
        flag(record, "ParentStopped")
        flag(record, "JobAssigned")
        flag(record, "Cancelled", False)
        flag(record, "StreamsComplete")
        flag(record, "StdoutTruncated", False)
        flag(record, "StderrTruncated", False)
        require(record.get("StartError") is None and record.get("StopError") is None and record.get("StreamError") is None,
                "probe: start/stream errors contradict success")
    return record


def runtime(record) -> dict:
    record = mapping(record, "selected runtime")
    path(record.get("Path"), "runtime.Path")
    require(record.get("Version") == "3.14.8" and record.get("PypdfVersion") == "6.19.0",
            "runtime: exact supported versions required")
    require(record.get("Source") in {"explicit", "application_venv", "py_launcher", "PATH"},
            "runtime: unknown selection provenance")
    require(record.get("Arguments") == ["-I", "-B"], "runtime: isolated engine flags required")
    origin = path(record.get("PypdfPath"), "runtime.PypdfPath")
    require(ntpath.basename(origin) == "__init__.py" and ntpath.basename(ntpath.dirname(origin)) == "pypdf",
            "runtime: unexpected pypdf package origin")
    executable_parent = ntpath.dirname(path(record["Path"], "runtime executable"))
    prefix = ntpath.dirname(executable_parent) if ntpath.basename(executable_parent) == "scripts" else executable_parent
    require(origin.startswith(ntpath.join(prefix, "lib", "site-packages", "pypdf") + "\\"),
            "runtime: pypdf is outside the selected interpreter's supported library")
    details = mapping(record.get("Details"), "runtime.Details")
    expected = {"protocol": "winbooksplit.runtime", "schema_version": 1,
                "version": "3.14.8", "implementation": "cpython", "platform": "win32",
                "machine": "AMD64", "bits": 64, "pypdf_version": "6.19.0"}
    for field, value in expected.items():
        require(type(details.get(field)) is type(value) and details[field] == value,
                f"runtime.Details.{field}: exact scalar evidence required")
    for field in ("ok", "isolated", "dont_write_bytecode"):
        flag(details, field)
    flag(details, "gil_disabled", False)
    same_path(details.get("executable"), record["Path"], "runtime: actual executable")
    same_path(details.get("pypdf_path"), record["PypdfPath"], "runtime: actual pypdf origin")
    text(details.get("message"), "runtime.Details.message")
    selected_probe = probe(record.get("Probe"), success=True)
    try:
        observed = json.loads(selected_probe["Stdout"])
    except (ValueError, TypeError) as error:
        raise ValueError("runtime: probe JSON missing") from error
    require(type(observed) is dict and observed == details, "runtime: reported details differ from native probe")
    attempts = record.get("Attempts")
    require(type(attempts) is list and bool(attempts), "runtime: selection attempts missing")
    last = mapping(attempts[-1], "runtime final attempt")
    flag(last, "Accepted")
    same_path(last.get("Path"), record["Path"], "runtime: accepted attempt")
    require(last.get("Source") == record["Source"] and last.get("Probe") == selected_probe,
            "runtime: selection provenance/probe differ")
    return record


def converter(record, calibre_path: str) -> dict:
    record = mapping(record, "selected converter")
    same_path(record.get("Path"), calibre_path, "converter: actual requested executable")
    require(record.get("Version") == "9.15.0" and record.get("Source") in {"explicit", "PATH", "known_location"},
            "converter: exact version and trusted provenance required")
    selected_probe = probe(record.get("Probe"), success=True)
    flag(selected_probe, "StderrTruncated", False)
    require(re.fullmatch(r"ebook-convert(?:\.exe)? \(calibre 9\.15\.0\)(?:\r?\nCreated by: Kovid Goyal <kovid@kovidgoyal\.net>)?(?:\r?\n)?",
                         selected_probe["Stdout"]) is not None,
            "converter: native version banner missing")
    attempts = record.get("Attempts")
    require(type(attempts) is list and bool(attempts), "converter: selection attempts missing")
    last = mapping(attempts[-1], "converter final attempt")
    flag(last, "Accepted")
    same_path(last.get("Path"), record["Path"], "converter: accepted attempt")
    require(last.get("Source") == record["Source"] and last.get("Probe") == selected_probe,
            "converter: selected provenance/probe differ")
    return record


def identity(record, label: str) -> dict:
    record = mapping(record, label)
    path(record.get("path"), label + ".path")
    digest(record.get("sha256"), label + ".sha256")
    integer(record.get("size_bytes"), label + ".size_bytes", 1)
    require(record.get("binding") == "reader_snapshot", label + ": captured reader binding required")
    return record


def writer_parity(case: dict) -> dict:
    writer = mapping(case.get("writer_result"), "writer_result")
    require(writer.get("mode") == "manual", "writer: actual manual case expected")
    total = integer(writer.get("total_pages"), "writer.total_pages", 1)
    source = identity(writer.get("source_identity"), "writer.source_identity")
    manifest = mapping(writer.get("manifest"), "writer.manifest")
    require(type(manifest.get("schema_version")) is int and manifest["schema_version"] == 1 and
            manifest.get("status") == "complete", "writer: complete manifest required")
    run_id = writer.get("run_id")
    require(type(run_id) is str and re.fullmatch(r"[0-9a-f]{32}", run_id), "writer: invalid run ID")
    final = path(writer.get("final_directory"), "writer.final_directory")
    require(ntpath.dirname(final) == path(case.get("output_base"), "case.output_base"),
            "writer: publication escaped requested base")
    require(total == 3 and len(writer.get("outputs", [])) == 3,
            "writer: three-page/three-chapter authored cases required")
    require(writer.get("manifest_filename") == "WinBookSplit_Manifest.json", "writer: manifest filename")
    for field in ("run_id", "final_directory", "mode", "total_pages", "source_identity",
                  "coverage", "written_count", "outputs"):
        require(field in writer and same_json(manifest.get(field), writer[field]), "manifest parity: " + field)
    entries = writer.get("outputs")
    outputs = case.get("outputs")
    require(type(entries) is list and bool(entries) and type(outputs) is list and len(outputs) == len(entries),
            "writer: output evidence missing")
    require(type(writer.get("written_count")) is int and writer["written_count"] == len(entries),
            "writer: written count mismatch")
    require(writer.get("coverage") == {"complete": True, "covered_pages": total, "section_count": len(entries)} and
            type(writer["coverage"].get("covered_pages")) is int and
            type(writer["coverage"].get("section_count")) is int and writer["coverage"].get("complete") is True,
            "writer: exact complete coverage required")
    names = set()
    start = 0
    page_ids = []
    for number, (entry, output) in enumerate(zip(entries, outputs), 1):
        entry = mapping(entry, "writer output")
        output = mapping(output, "observed output")
        require(type(entry.get("sequence")) is int and entry["sequence"] == number, "writer: sequence mismatch")
        require(type(entry.get("start")) is int and entry["start"] == start, "writer: gap/overlap")
        end = integer(entry.get("end"), "writer.end", start + 1)
        require(end <= total and type(entry.get("page_count")) is int and entry["page_count"] == end - start,
                "writer: invalid page range")
        name = text(entry.get("filename"), "writer.filename")
        require(ntpath.basename(name) == name and not any(char in name for char in '<>:"/\\|?*\x00') and
                name.lower().endswith(".pdf") and name.casefold() not in names, "writer: unsafe/duplicate filename")
        names.add(name.casefold())
        digest(entry.get("sha256"), "writer output SHA256")
        integer(entry.get("size_bytes"), "writer output size", 1)
        require(output.get("filename") == name and output.get("range") == [start, end] and
                output.get("page_ids") == list(range(start + 1, end + 1)) and
                output.get("sha256") == entry["sha256"],
                "writer: observed output identity/range mismatch")
        require(type(output.get("range")) is list and all(type(v) is int for v in output["range"]) and
                type(output.get("page_ids")) is list and all(type(v) is int for v in output["page_ids"]),
                "writer: physical ranges/IDs require integers")
        if "size_bytes" in output:
            require(type(output["size_bytes"]) is int and output["size_bytes"] == entry["size_bytes"],
                    "writer: observed size mismatch")
        if "page_count" in output:
            require(type(output["page_count"]) is int and output["page_count"] == end - start,
                    "writer: observed page count mismatch")
        page_ids.extend(output["page_ids"])
        start = end
    require(start == total and page_ids == list(range(1, total + 1)), "writer: physical page omission/duplication")
    for field in ("independent_source_page_content_sha256", "page_content_sha256"):
        vector = case.get(field)
        require(type(vector) is list and len(vector) == total, field + ": missing page content")
        for value in vector:
            digest(value, field)
    require(case["independent_source_page_content_sha256"] == case["page_content_sha256"], "writer: page content mismatch")
    flag(case, "manifest_validated")
    flag(case, "owned_outputs_removed")
    originals = mapping(case.get("input_and_neighbor_identity"), "source/neighbor identities")
    input_identity = originals.get(case.get("source_path"))
    input_identity = mapping(input_identity, "original input identity")
    original = writer.get("original_ebook_identity")
    require((original or source).get("sha256") == input_identity.get("sha256"), "writer: input identity mismatch")
    same_path((original or source).get("path"), case["source_path"], "writer: original source path")
    return writer


def streams(record: dict, label: str) -> None:
    for field in ("stdout", "stderr"):
        require(type(record.get(field)) is str, label + ": actual streams missing")
    require(type(record.get("exit_code")) is int, label + ": native status missing")


def policies(value, label: str) -> None:
    require(type(value) is list and len(value) == 4, label + ": four stored scopes required")
    require({mapping(item, label).get("scope") for item in value} ==
            {"MachinePolicy", "UserPolicy", "CurrentUser", "LocalMachine"}, label + ": scopes differ")
    for item in value:
        text(item.get("policy"), label + ".policy")


def common_case(case: dict, hosts: dict[str, dict], source_map: dict) -> None:
    host = "PS51" if case["id"].startswith("BAT-") else case["id"].split("-")[0]
    same_path(case.get("shell_executable"), hosts[host]["shell_executable"], "case: requested host")
    command = case.get("command")
    require(type(command) is list and all(type(value) is str for value in command), "case: actual argv missing")
    if case["id"].startswith("BAT-"):
        require(len(command) == 5 and ntpath.basename(path(command[0], "BAT executable")) == "cmd.exe" and
                command[1:4] == ["/d", "/v:off", "/c"], "BAT: actual fixed command/no CALL scope required")
    else:
        require(len(command) == 6 and command[1:5] == ["-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File"],
                "case: actual supported-host argv required")
        same_path(command[0], hosts[host]["shell_executable"], "case: actual host command")
    path(command[-1], "case wrapper")
    path(case.get("cwd"), "case.cwd")
    path(case.get("output_base"), "case.output_base")
    source_path = path(case.get("source_path"), "case.source_path")
    require(case.get("source_format") in {"pdf", "epub"} and
            ntpath.splitext(source_path)[1] == "." + case["source_format"], "case: source format mislabeled")
    for field in ("source_read_only_attribute_observed", "source_attributes_restored",
                  "input_unchanged", "neighbor_unchanged", "decoys_unchanged",
                  "decoy_execution_marker_absent", "module_import_marker_absent",
                  "copied_application_unchanged", "owned_outputs_removed"):
        flag(case, field)
    originals = mapping(case.get("input_and_neighbor_identity"), "original identities")
    require(len(originals) == 2 and case["source_path"] in originals, "case: source/neighbor identities missing")
    require(case.get("input_and_neighbor_identity_after") == originals, "case: source/neighbor changed")
    for filename, observed in originals.items():
        path(filename, "original path")
        observed = mapping(observed, "original identity")
        digest(observed.get("sha256"), "original SHA256")
        integer(observed.get("size_bytes"), "original size", 1)
        for field in ("device", "inode"):
            integer(observed.get(field), "original." + field, 1)
        integer(observed.get("attributes"), "original.attributes")
    application = mapping(case.get("application_path_sha256"), "copied source hashes")
    expected_files = {"VERSION", "WinBookSplit.ps1", "WinBookSplit.bat", "requirements.txt",
                      "engine/WinBookSplit.Paths.ps1", "engine/WinBookSplit.Diagnostics.ps1",
                      "engine/WinBookSplit.Runtime.ps1", "engine/winbooksplit_engine.py",
                      "engine/WinBookSplit.Process.ps1", "engine/WinBookSplit.Outcomes.json",
                      "engine/winbooksplit_windows.py", "engine/winbooksplit_conversion.py",
                      "engine/winbooksplit_job.py", "engine/WinBookSplit.Logging.ps1",
                      "engine/WinBookSplit.Support.ps1", "Export-WinBookSplitDiagnostics.ps1"}
    require(set(application) == expected_files and
            all(application[name] == source_map.get(name) for name in expected_files),
            "case: copied application differs from tested bytes")
    decoys = mapping(case.get("decoy_sha256"), "decoy hashes")
    require(case.get("decoy_sha256_after") == decoys, "case: decoy bytes changed")
    for filename, value in decoys.items():
        path(filename, "decoy path")
        digest(value, "decoy SHA256")
    markers = mapping(case.get("decoy_marker_observation"), "decoy observations")
    for field in ("native", "module", "converter_version", "conversion"):
        path(markers.get(field + "_path"), "decoy marker path")
        require(type(markers.get(field + "_absent")) is bool, "decoy marker observation missing")
    flag(markers, "native_absent")
    flag(markers, "module_absent")
    flag(markers, "conversion_absent")
    require(type(case.get("controlled_PATH")) is str and bool(case["controlled_PATH"]), "case: controlled PATH missing")
    require(case.get("controlled_PYTHONPATH") is None or type(case["controlled_PYTHONPATH"]) is str,
            "case: controlled PYTHONPATH type")
    streams(case, case["id"])


def failure_case(case: dict) -> None:
    require(case["exit_code"] == 3, "failure: dependency native status 3 required")
    for field in ("preflight_before_output", "no_success_summary", "setup_guidance_verified"):
        flag(case, field)
    for field in ("engine_invocation_record_absent", "dependency_success_record_absent"):
        flag(case, field)
    flag(case, "conversion_started", False)
    require(case.get("output_base_before") == [] and case.get("output_base_after") == [] and
            type(case.get("written_count")) is int and case["written_count"] == 0 and case.get("outputs") == [],
            "failure: pre-write empty base/zero chapters required")
    require("dependency_selection" not in case and "engine_invocation" not in case and "writer_result" not in case,
            "failure: engine execution contradicts preflight rejection")
    error = mapping(case.get("dependency_error"), "dependency error")
    expected_code = "runtime_invalid" if case["category"] == "python_failure" else (
        "converter_invalid" if case["kind"] == "wrong-version" else "converter_not_found")
    require(error.get("Code") == expected_code, "failure: wrong dependency category")
    try:
        frames = [json.loads(line.removeprefix("[DEPENDENCY-ERROR] ")) for line in case["stdout"].splitlines()
                  if line.startswith("[DEPENDENCY-ERROR] ")]
    except ValueError as error_parse:
        raise ValueError("failure: malformed actual error frame") from error_parse
    require(frames == [error] and "Dependency preflight failed" in case["stdout"] and
            not any(fragment in case["stdout"] for fragment in ("[ENGINE]", "[DEPENDENCY] ", "Done.", "[Writing]", "[CONVERSION]")),
            "failure: actual rejection frame/success contradiction")
    attempts = error.get("Attempts")
    require(type(attempts) is list, "failure: rejection attempts missing")
    for attempt in attempts:
        attempt = mapping(attempt, "rejected attempt")
        flag(attempt, "Accepted", False)
        path(attempt.get("Path"), "rejected candidate")
        text(attempt.get("Source"), "rejected source")
        text(attempt.get("Message"), "rejection reason")
        if attempt.get("Probe") is not None:
            probe(attempt["Probe"], success=False)
    explicit = case["category"] == "python_failure" or case["kind"] == "wrong-version"
    if explicit:
        require(len(attempts) == 1 and attempts[0]["Source"] == "explicit", "failure: explicit path silently fell back")
        same_path(case.get("candidate_path"), attempts[0]["Path"], "failure: attempted explicit path")
    else:
        require(case.get("candidate_path") is None, "failure: missing converter invented selected path")
    if case["category"] == "python_failure":
        flag(case, "converter_version_marker_absent")
        flag(case["decoy_marker_observation"], "converter_version_absent")
        require(all(fragment in case["stdout"] for fragment in
                    ("3.14.8", "6.19.0", "-PythonPath", "-I -m pip", "requirements.txt", case["candidate_path"])),
                "failure: exact-interpreter setup guidance missing")
        if case["kind"] != "missing":
            require(attempts[0].get("Probe") is not None, "failure: actual incompatible candidate probe missing")
            actual = attempts[0]["Probe"]
            flag(actual, "ParentStopped")
            flag(actual, "StreamsComplete")
            flag(actual, "TimedOut", False)
            require(type(actual.get("ExitCode")) is int, "failure: actual incompatible status missing")
            if case["kind"] == "non-python":
                require(actual["ExitCode"] == 7 and "OWNED-NON-PYTHON-NATIVE-CONTROL" in actual["Stdout"],
                        "failure: actual non-Python control mislabeled")
            else:
                try:
                    details = json.loads(actual["Stdout"])
                except ValueError as malformed:
                    raise ValueError("failure: actual Python rejection record missing") from malformed
                require(type(details) is dict and details.get("protocol") == "winbooksplit.runtime" and
                        type(details.get("schema_version")) is int and details["schema_version"] == 1 and
                        details.get("ok") is False and actual["ExitCode"] == 1,
                        "failure: actual Python rejection protocol/status differs")
                message = text(details.get("message"), "actual Python rejection reason")
                if case["kind"] == "wrong-version":
                    require(details.get("version") == "3.14.7" and "3.14.8" in message,
                            "failure: real unsupported Python control mislabeled")
                elif case["kind"] == "missing-pypdf":
                    require(details.get("version") == "3.14.8" and "No module named 'pypdf'" in message,
                            "failure: missing-pypdf control mislabeled")
                else:
                    require(details.get("version") == "3.14.8" and details.get("pypdf_version") == "9.99.0" and
                            "pypdf 6.19.0" in message, "failure: altered owned pypdf control mislabeled")
        else:
            require(attempts[0].get("Probe") is None, "failure: missing interpreter was executed")
    else:
        require("Calibre 9.15.0" in case["stdout"] and "-CalibrePath" in case["stdout"],
                "failure: converter setup guidance missing")
        if case["kind"] == "wrong-version":
            require(attempts[0].get("Probe") is not None and attempts[0]["Probe"]["ExitCode"] == 0 and
                    "9.14.0" in attempts[0]["Probe"]["Stdout"], "failure: actual wrong converter version missing")
            flag(case["decoy_marker_observation"], "converter_version_absent", False)


def success_case(case: dict, calibre_path: str) -> None:
    require(case["exit_code"] == 0, "success: actual native zero required")
    flag(case, "selected_interpreter_used_for_engine")
    flag(case, "all_physical_pages_preserved")
    selected = mapping(case.get("dependency_selection"), "dependency selection")
    selected_runtime = runtime(selected.get("Runtime"))
    same_path(selected_runtime["Path"], case.get("expected_python_path"), "success: expected interpreter")
    expected_source = {"explicit-literal": "explicit", "app-venv": "application_venv",
                       "py-launcher-listing": "py_launcher"}.get(case["kind"], "PATH")
    require(case.get("expected_python_source") == selected_runtime["Source"] == expected_source,
            "success: interpreter priority/provenance differs")
    invocation = mapping(case.get("engine_invocation"), "actual engine invocation")
    same_path(invocation.get("path"), selected_runtime["Path"], "success: same interpreter used")
    argv = invocation.get("arguments")
    require(type(argv) is list and all(type(value) is str for value in argv) and argv[:4] == ["-I", "-B", "-X", "utf8"],
            "success: actual isolated UTF-8 engine argv missing")
    require(len(argv) >= 11 and argv[7:9] == ["manual", ""], "success: authored prompted manual control differs")
    same_path(argv[4], ntpath.join(ntpath.dirname(case["command"][-1]), "a", "engine", "winbooksplit_engine.py"),
              "success: shipped copied engine")
    same_path(argv[5], case["source_path"], "success: engine source")
    same_path(argv[6], case["output_base"], "success: engine output base")
    origin = path(selected_runtime["PypdfPath"], "selected pypdf")
    for excluded in (ntpath.dirname(path(case["source_path"], "source")), path(case["cwd"], "cwd")):
        require(not origin.startswith(excluded + "\\"), "success: shadowed pypdf origin")
    writer = writer_parity(case)
    require(case.get("interaction", {}).get("engine_invocations") == [invocation], "Runtime interaction used a different invocation")
    interactions.validate(case.get("interaction"), case.get("process_summaries"), ["input_ready", "plan_ready"],
                          ["starts", "execute"], starts="2,3", execution=writer)
    require(case.get("stdin_utf8") == ("M\nN\n2,3\nY\nN\n\n" if case["source_format"] != "pdf" else "M\n2,3\nY\nN\n\n"),
            "Runtime actual explicit confirmation answers differ")
    frame = mapping(case.get("engine_record"), "actual engine result")
    evidence = case.get("console_evidence")
    console.validate_console_evidence(evidence, evidence.get("operation_outcome") if isinstance(evidence, dict) else None,
        evidence.get("log_sha256") if isinstance(evidence, dict) else None)
    require(evidence["operation_outcome"].get("engine_result") == frame, "Runtime run record differs from actual engine frame")
    require(frame.get("protocol") == "winbooksplit.result" and type(frame.get("version")) is int and
            frame["version"] == 1 and frame.get("status") == "success" and frame.get("code") == "split_complete" and
            type(frame.get("exit_code")) is int and frame["exit_code"] == 0 and frame.get("mode") == "manual" and
            same_json(frame.get("execution"), writer) and type(frame.get("written_count")) is int and
            frame["written_count"] == case.get("written_count") == writer["written_count"],
            "success: actual protocol/writer/count parity differs")
    if case["source_format"] == "pdf":
        require("Converter" in selected and selected["Converter"] is None and len(argv) == 11,
                "PDF: converter must remain unprobed")
        flag(case["decoy_marker_observation"], "converter_version_absent")
    else:
        chosen = converter(selected.get("Converter"), calibre_path)
        converter_source = "explicit" if case["kind"] == "explicit-portable" else "PATH"
        require(case.get("expected_converter_source") == chosen["Source"] == converter_source,
                "success: converter selection priority differs")
        require(len(argv) == 15 and argv[11] == "--calibre-path" and argv[13:] == ["--conversion-timeout", "1800"],
                "success: engine conversion argv differs")
        same_path(argv[12], chosen["Path"], "success: same converter used")
        original = mapping(writer.get("original_ebook_identity"), "original ebook identity")
        require(original.get("binding") == "ebook_snapshot", "success: original ebook binding missing")
        conversion = mapping(writer.get("conversion"), "actual conversion")
        require(conversion.get("original_source_identity") == original and
                conversion.get("workspace_cleanup") == {"cleanup_complete": True, "retained_staging": None},
                "success: original/conversion workspace parity")
        generated = mapping(conversion.get("generated_pdf_identity"), "generated identity")
        require(generated.get("sha256") == writer["source_identity"]["sha256"] and generated.get("page_count") == 3 and
                generated.get("page_content_sha256") == case["page_content_sha256"] and
                writer.get("retained_intermediate", "missing") is None,
                "success: actual captured conversion evidence missing")
        same_path(conversion.get("converter_path"), chosen["Path"], "success: converted with selected path")
        require(type(conversion.get("exit_code")) is int and conversion["exit_code"] == 0,
                "success: actual converter return code missing")
        for field in ("original_ebook_identity", "conversion", "retained_intermediate"):
            require(writer["manifest"].get(field) == writer[field], "ebook manifest parity: " + field)
    if case["kind"] == "multiple-path":
        rejected = [a for a in selected_runtime["Attempts"] if a.get("Source") == "PATH" and a.get("Accepted") is False]
        require(bool(rejected) and any(a.get("Probe") and "3.14.7" in a["Probe"]["Stdout"] for a in rejected),
                "success: earlier incompatible PATH interpreter not observed")
    if case["category"] == "shadow":
        excluded = {path(name, "shadow executable") for name in case["decoy_sha256"] if name.endswith(".exe")}
        observations = {path(a.get("Path"), "excluded shadow") for a in selected_runtime["Attempts"]
                        if a.get("Accepted") is False and a.get("Probe") is None}
        require(excluded <= observations, "shadow: automatic exclusion observations missing")
    if case["kind"] == "py-launcher-listing":
        listings = [a for a in selected_runtime["Attempts"] if a.get("Source") == "py_launcher_listing"]
        require(len(listings) == 1 and listings[0].get("Accepted") is True and
                ntpath.basename(path(listings[0].get("Path"), "launcher path")) == "py.exe",
                "success: actual launcher listing attempt missing")
        listing = probe(listings[0].get("Probe"), success=True)
        require(selected_runtime["Path"] in listing["Stdout"] + listing["Stderr"],
                "success: launcher did not list selected interpreter")
        control = mapping(case.get("launcher_listing_control"), "launcher listing control")
        require(control.get("scope") == "authored-native-py-listing-control; actual selected interpreter" and
                control.get("arguments") == ["-0p"] and control.get("control_reached") is True,
                "success: authored native read-only listing control missing")
        same_path(control.get("path"), listings[0]["Path"], "launcher control executable")
        same_path(control.get("listed_python_path"), selected_runtime["Path"], "launcher listed actual Python")
        path(control.get("marker_path"), "launcher control marker")
        digest(control.get("marker_sha256"), "launcher control marker hash")
        digest(control.get("sha256"), "launcher control executable hash")


def validate_runtime_report(child, requested_shells, calibre_path) -> None:
    child = mapping(child, "runtime child")
    require(type(child.get("schema_version")) is int and child["schema_version"] == 1 and
            child.get("task_id") == "M2-T04" and child.get("result") == "RUNTIME_REGRESSION_PASSED" and
            child.get("success") is True and type(child.get("exit_code")) is int and child["exit_code"] == 0 and
            child.get("acceptance_ids") == ACCEPTANCE, "runtime: successful exact task/acceptance envelope required")
    for field in ("source_unchanged", "baseline_guards_preserved", "input_and_neighbor_unchanged",
                  "machine_settings_unchanged", "owned_temp_removed"):
        flag(child, field)
    require(child.get("immutable_original_commit") == ORIGINAL_COMMIT, "runtime: immutable historical guard identity")
    source_map = mapping(child.get("tested_path_sha256"), "tested source map")
    required_source = {"VERSION", "WinBookSplit.ps1", "WinBookSplit.bat", "requirements.txt",
                       "engine/WinBookSplit.Runtime.ps1", "tests/runtime/characterize_runtime.py",
                       "tests/runtime/validate_runtime_report.py", "tests/runtime/README.md",
                       "tests/python/test_runtime_receipt.py", "tests/baseline/expected_original.json",
                       "tests/extraction/characterize_extraction.py", "tests/fixtures/generate_pdf_fixtures.py"}
    require(required_source <= source_map.keys(), "runtime: tested implementation/harness/historical map incomplete")
    for filename, value in source_map.items():
        text(filename, "tested source filename")
        digest(value, "tested source hash")
    require(type(requested_shells) is list and len(requested_shells) == 2 and
            len({path(p, "requested shell") for p in requested_shells}) == 2,
            "runtime: two distinct requested supported hosts required")
    host_cases = child.get("host_cases")
    require(type(host_cases) is list and len(host_cases) == 2 and
            {mapping(h, "host").get("id") for h in host_cases} == set(HOSTS), "runtime: both actual host records required")
    hosts = {host["id"]: host for host in host_cases}
    require({path(h.get("shell_executable"), "host executable") for h in host_cases} ==
            {path(p, "requested host") for p in requested_shells}, "runtime: unrequested/reused host")
    for identifier, host in hosts.items():
        flag(host, "passed")
        streams(host, "host")
        require(host["exit_code"] == 0 and type(host.get("host_major")) is int and
                host["host_major"] == (5 if identifier == "PS51" else 7), "host: actual version/status missing")
        version = text(host.get("host_version"), "host version")
        require(version.startswith("5.1.") if identifier == "PS51" else version == "7.6.5", "host: mislabeled version")
        require(type(host.get("syntax_error_count")) is int and host["syntax_error_count"] == 0 and
                host.get("syntax_checked") == ["WinBookSplit.ps1", "engine/WinBookSplit.Paths.ps1",
                    "engine/WinBookSplit.Diagnostics.ps1", "engine/WinBookSplit.Runtime.ps1", "engine/WinBookSplit.Process.ps1"], "host: required syntax observations missing")
        policies(host.get("stored_policies"), "host stored policy")
        require(host.get("policies_after") == host["stored_policies"], "host: stored policies changed")
        command = host.get("command")
        require(type(command) is list and len(command) == 7 and
                command[1:6] == ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-File"],
                "host: actual observation command missing")
        same_path(command[0], host["shell_executable"], "host: observed executable")
        path(command[-1], "host syntax probe")
        path(host.get("cwd"), "host observation cwd")
        observed = json.loads(host["stdout"])
        require(type(observed) is dict and all(observed.get(k) == host.get(k) for k in
                ("host_major", "host_version", "syntax_error_count", "syntax_checked", "stored_policies")),
                "host: actual observation frame differs")
    calibre = mapping(child.get("calibre"), "actual Calibre")
    same_path(calibre.get("path"), calibre_path, "Calibre requested path")
    require(calibre.get("version") == "9.15.0" and
            calibre.get("sha256") == "f46a01c9b8cd392e1190ce2e50a2a5a48b326c8c3eab670bbcd8b298bede501d",
            "runtime: pinned actual Calibre identity required")
    groups = {field: case_set(child, field) for field in CASE_KINDS}
    require(type(child.get("case_count")) is int and child["case_count"] == sum(len(g) for g in groups.values()) == 33,
            "runtime: all thirty-three cases required")
    for field, cases in groups.items():
        for case in cases.values():
            common_case(case, hosts, source_map)
            if field.endswith("failure_cases"):
                failure_case(case)
            else:
                success_case(case, calibre_path)
            decoys = case["decoy_sha256"]
            names = [ntpath.basename(filename) for filename in decoys]
            if field == "shadow_cases":
                require(len(names) == 7 and names.count("pypdf.py") == 3 and
                        names.count("python.exe") == names.count("py.exe") == 2 and
                        bool(case["controlled_PYTHONPATH"]), "shadow: concrete book/CWD/environment decoys missing")
            elif case["category"] == "converter_failure" and case["kind"] == "book-cwd-decoys":
                require(names == ["ebook-convert.exe", "ebook-convert.exe"],
                        "converter: concrete book/CWD decoys missing")
                require(len(case["dependency_error"]["Attempts"]) == 2 and
                        all(a.get("Probe") is None for a in case["dependency_error"]["Attempts"]),
                        "converter: decoy was probed or rejection observations missing")
                require({path(a["Path"], "excluded converter") for a in case["dependency_error"]["Attempts"]} ==
                        {path(name, "converter decoy") for name in decoys}, "converter: exclusion paths differ")
            else:
                require(not decoys, "runtime: unexpected decoy case scope")
    controls = mapping(child.get("positive_decoy_controls"), "positive decoy controls")
    for name in ("native", "module"):
        control = mapping(controls.get(name), "positive " + name)
        flag(control, "control_reached")
        streams(control, "positive control")
        require(control["exit_code"] == 0 if name == "native" else control["exit_code"] != 0,
                "positive control: actual status differs")
        path(control.get("marker_path"), "positive marker")
        digest(control.get("marker_sha256"), "positive marker hash")
        path(control.get("cwd"), "positive cwd")
        require(type(control.get("command")) is list and bool(control["command"]), "positive actual argv missing")
    require(controls["native"]["command"][1:] == ["--version"], "native marker control argv")
    require(controls["module"]["command"][1:] == ["-B", "-c", "import pypdf"], "module marker control argv")
    digest(controls["module"].get("module_sha256"), "positive module hash")
    path(controls["module"].get("module_path"), "positive module path")
    path(controls["module"].get("controlled_PYTHONPATH"), "positive module PYTHONPATH")
    compilation = mapping(child.get("native_control_compilation"), "native control compilation")
    streams(compilation, "native compilation")
    require(compilation["exit_code"] == 0 and type(compilation.get("command")) is list and
            bool(compilation["command"]), "native control: actual compile failed")
    digest(compilation.get("source_sha256"), "native control source")
    digest(compilation.get("executable_sha256"), "native control executable")
    unsupported = mapping(child.get("unsupported_real_python_observation"), "unsupported real Python")
    streams(unsupported, "unsupported real Python")
    require(unsupported["exit_code"] == 0 and "3.14.7" in unsupported["stdout"],
            "runtime: real unsupported interpreter control missing")
    reference = mapping(child.get("independent_real_conversion_reference"), "independent conversion")
    reference_process = mapping(reference.get("process"), "independent converter process")
    require(reference_process.get("exit_code") == 0 and type(reference_process.get("argv")) is list and
            reference_process["argv"][-2:] == ["--output-profile", "tablet"], "runtime: actual reference conversion missing")
    same_path(reference_process["argv"][0], calibre_path, "reference converter path")
    require(type(reference.get("page_count")) is int and reference["page_count"] == 3 and
            type(reference.get("page_content_sha256")) is list and len(reference["page_content_sha256"]) == 3,
            "runtime: three-page independent conversion reference missing")
    digest(reference.get("sha256"), "reference PDF SHA256")
    integer(reference.get("size_bytes"), "reference PDF size", 1)
    for value in reference["page_content_sha256"]:
        digest(value, "reference page content")
    for case in groups["converter_success_cases"].values():
        require(case["independent_source_page_content_sha256"] == reference["page_content_sha256"],
                "runtime: real conversion differs from independent physical reference")
