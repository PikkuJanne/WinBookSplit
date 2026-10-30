"""Finite final-ZIP native acceptance oracle; stdlib plus the locked pypdf only.

Native processes belong to Invoke-NativeAcceptance.ps1 and the shipped Process
helper. This module never starts a process or deletes an artifact. Console
receipt helpers below retain the existing current_launchers.py checks; physical
publication checks retain characterize_manual/conversion's reopened-page oracle.
"""
from __future__ import annotations

import argparse
from datetime import datetime, timezone
from hashlib import sha256
import importlib.util
import json
import math
import os
from pathlib import Path
import re
import stat
import sys
from types import SimpleNamespace

MAX_JSON = 16 * 1024 * 1024
NO_BOOKMARK_EXIT = 5
NEGATIVE_CODES = {"preview": "preview_complete", "invalid-manual": "invalid_arguments",
                  "no-bookmarks": "no_bookmarks", "cancel-plan": "processing_cancelled",
                  "missing-python": "runtime_invalid", "missing-pypdf": "setup_required"}
CASE_IDS = ("ref-epub", "ref-azw3", "manual", "manual-first", "auto1", "auto2",
            "repeat", "preview", "invalid-manual", "no-bookmarks", "cancel-plan",
            "missing-python", "missing-pypdf", "epub-discard", "epub-keep",
            "azw3-discard", "azw3-keep", "bat-one", "bat-two")
APP_VERSION = None
RUNTIME_POLICY = None
RUNTIME_POLICY_SHA256 = None
compatibility = None


class OracleError(ValueError):
    """Fixed identifier safe to include in the redacted return report."""


def need(condition, code):
    if not condition:
        raise OracleError(code)


require = need


def windows_system_directory():
    """Use the native known-system path, preserving its literal returned casing."""
    import ctypes
    from ctypes import wintypes
    need(os.name == "nt", "native_windows_required")
    query = ctypes.WinDLL("kernel32", use_last_error=True).GetSystemDirectoryW
    query.argtypes = (wintypes.LPWSTR, wintypes.UINT)
    query.restype = wintypes.UINT
    buffer = ctypes.create_unicode_buffer(32768)
    length = query(buffer, len(buffer))
    need(0 < length < len(buffer) and re.fullmatch(r"[A-Za-z]:\\[^\x00-\x1f]+", buffer.value), "native_system_directory_invalid")
    return str(ordinary(Path(buffer.value), directory=True))


def is_reparse(path):
    return bool(path.lstat().st_file_attributes & stat.FILE_ATTRIBUTE_REPARSE_POINT) if os.name == "nt" else path.is_symlink()


history = SimpleNamespace(is_reparse=is_reparse)


def ordinary(path, *, directory=False):
    path = Path(path)
    need(path.is_absolute() and path.resolve(strict=True) == path, "ordinary_path_required")
    for ancestor in [*reversed(path.parents), path]:
        details = ancestor.lstat()
        need(not is_reparse(ancestor), "reparse_path_rejected")
        need(stat.S_ISDIR(details.st_mode) if ancestor != path or directory else stat.S_ISREG(details.st_mode), "ordinary_kind_required")
    if not directory:
        need(path.stat().st_nlink == 1, "hardlink_file_rejected")
    return path


def unique_json(raw):
    def pairs(items):
        result = {}
        for key, value in items:
            need(key not in result, "duplicate_json_key")
            result[key] = value
        return result
    return json.loads(raw, object_pairs_hook=pairs, parse_constant=lambda value: (_ for _ in ()).throw(OracleError("nonfinite_json")))


def read_json(path):
    path = ordinary(path)
    need(0 < path.stat().st_size <= MAX_JSON, "json_size_invalid")
    return unique_json(path.read_bytes().decode("utf-8", errors="strict"))


def write_json(path, value):
    path = Path(path)
    ordinary(path.parent, directory=True)
    need(not path.exists() and not path.is_symlink(), "output_already_exists")
    raw = (json.dumps(value, ensure_ascii=True, allow_nan=False, indent=2) + "\n").encode("utf-8")
    need(len(raw) <= MAX_JSON, "json_size_invalid")
    with path.open("xb") as stream:
        stream.write(raw)


def observation(path):
    path = ordinary(path)
    details = path.stat()
    return {"path": str(path), "sha256": sha256(path.read_bytes()).hexdigest(), "size_bytes": details.st_size,
            "device": details.st_dev, "inode": details.st_ino}


def unchanged(item):
    need(observation(Path(item["path"])) == item, "immutable_file_changed")


def bind_application(root, policy_digest):
    global APP_VERSION, RUNTIME_POLICY, RUNTIME_POLICY_SHA256, compatibility
    root = ordinary(root, directory=True)
    APP_VERSION = ordinary(root / "VERSION").read_text(encoding="utf-8").strip()
    need(re.fullmatch(r"\d+\.\d+\.\d+", APP_VERSION), "application_version_invalid")
    helper = ordinary(root / "engine/winbooksplit_compatibility.py")
    spec = importlib.util.spec_from_file_location("wbs_native_compatibility", helper)
    compatibility = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(compatibility)
    policy_path = ordinary(root / "engine/WinBookSplit.Compatibility.json")
    RUNTIME_POLICY, raw = compatibility._read_policy(policy_path)
    RUNTIME_POLICY_SHA256 = sha256(raw).hexdigest()
    need(RUNTIME_POLICY_SHA256 == policy_digest, "policy_pin_mismatch")
    return root


def python_version_supported(value):
    if type(value) is not str or re.fullmatch(r"\d+\.\d+\.\d+", value) is None:
        return False
    record = {"version": value, "version_info": [int(part) for part in value.split(".")], "releaselevel": "final", "serial": 0,
              "implementation": "cpython", "platform": "win32", "machine": "AMD64", "bits": 64, "gil_disabled": False,
              "isolated": True, "dont_write_bytecode": True, "executable": "C:\\version-only\\python.exe"}
    return compatibility.validate_interpreter(record, RUNTIME_POLICY)["ok"]


# Minimal reviewed console closure from tests/manual/current_launchers.py
# Raw source SHA256: dc3950e277ed367414a56a52aadeafd0dc44d2fdb31fe5d6fd61de2c7c7af465
# Only adaptations: actual VERSION and policy-derived converter lock.
CONSOLE_MEMBERS = {".WinBookSplit-console-owner.json", "console.log", "WinBookSplit_Run.json"}

def _console_need(condition, message):
    if not condition:
        raise ValueError(message)

def _unique_json(raw):
    def pairs(items):
        result = {}
        for key, value in items:
            _console_need(key not in result, "Duplicate local run-record field")
            result[key] = value
        return result
    return json.loads(raw, object_pairs_hook=pairs)

def confirmed_console_plan(text):
    """A displayed plan becomes confirmed only through its exact execute reply."""
    plans = [_unique_json(line[len("[PLAN] "):]) for line in text.splitlines() if line.startswith("[PLAN] ")]
    replies = [_unique_json(line[len("[INTERACTION-REPLY] "):]) for line in text.splitlines() if line.startswith("[INTERACTION-REPLY] ")]
    execute = [reply for reply in replies if reply.get("action") == "execute"]
    if not execute:
        return None
    _console_need(len(execute) == 1, "Local console has conflicting execute confirmations")
    reply = execute[0]
    matches = [event for event in plans if all(event.get(key) == reply.get(key) for key in ("protocol", "version", "session", "sequence"))
               and event.get("stage") == "plan_ready" and event.get("plan_sha256") == reply.get("plan_sha256")]
    _console_need(len(matches) == 1 and reply.get("protocol") == "winbooksplit.interaction" and type(reply.get("version")) is int
        and reply["version"] == 1 and type(reply.get("sequence")) is int and reply["sequence"] > 0
        and re.fullmatch(r"[0-9a-f]{32}", reply.get("session", "")), "Local console execute reply does not bind one displayed plan")
    event = matches[0]
    canonical = event.get("plan_json")
    _console_need(isinstance(canonical, str) and sha256(canonical.encode("utf-8")).hexdigest() == event["plan_sha256"]
        and _unique_json(canonical) == event.get("plan"), "Local console confirmed plan digest/data differ")
    return event["plan"]

def validate_console_evidence(receipt, outcome, log_sha256, *, source_path=None, source_observation=None, allow_unfinalized=False, log_text=None, expected_python_version=None):
    """Pure receipt checks; actual native bytes/held cleanup are separate proofs."""
    need = _console_need
    need(isinstance(receipt, dict) and isinstance(outcome, dict) and receipt.get("log_sha256") == log_sha256
         and re.fullmatch(r"[0-9a-f]{64}", log_sha256 or ""), "Local console receipt/log hash missing")
    need(outcome.get("protocol") == "winbooksplit.outcome" and type(outcome.get("version")) is int and outcome["version"] == 1
         and type(outcome.get("exit_code")) is int and outcome["exit_code"] in {0, 2, 3, 4, 5, 6, 7, 130}
         and type(outcome.get("written_count")) is int and outcome["written_count"] >= 0, "Local console terminal outcome invalid")
    owner, identities = receipt.get("owner"), receipt.get("file_identities")
    directory_identity = receipt.get("directory_identity")
    need(isinstance(directory_identity, dict) and set(directory_identity) == {"device", "inode"}
         and type(directory_identity["device"]) is int and directory_identity["device"] >= 0
         and type(directory_identity["inode"]) is int and directory_identity["inode"] > 0, "Local console directory identity missing")
    need(isinstance(owner, dict) and set(owner) == {"run_id", "kind"} and owner["kind"] == "console"
         and re.fullmatch(r"[0-9a-f]{32}", owner.get("run_id", "")), "Local console owner invalid")
    state = receipt.get("state")
    members = CONSOLE_MEMBERS if state == "finalized" else CONSOLE_MEMBERS - {"WinBookSplit_Run.json"}
    need(receipt.get("members") == sorted(members) and isinstance(identities, dict) and set(identities) == members,
         "Local console members/identities differ")
    for item in identities.values():
        need(isinstance(item, dict) and all(type(item.get(key)) is int and item[key] >= 0 for key in ("device", "inode", "size_bytes"))
             and item["inode"] > 0 and re.fullmatch(r"[0-9a-f]{64}", item.get("sha256", "")), "Local console file identity invalid")
    need(identities["console.log"]["sha256"] == log_sha256 and identities["console.log"]["size_bytes"] == receipt.get("log_size_bytes"),
         "Local console actual length/hash differ")
    if log_text is not None:
        need(isinstance(log_text, str) and sha256(log_text.encode("utf-8")).hexdigest() == log_sha256
            and len(log_text.encode("utf-8")) == receipt["log_size_bytes"]
            and receipt.get("confirmed_plan") == confirmed_console_plan(log_text), "Local console raw consent/plan binding differs")
    if state == "unfinalized":
        need(allow_unfinalized and outcome.get("exit_code") == 6 and outcome.get("status") == "incomplete"
             and receipt.get("run_manifest") is None and receipt.get("run_manifest_raw") is None
             and receipt.get("run_manifest_sha256") is None, "Unfinalized console is not the declared finalizer failure")
        return
    need(state == "finalized" and receipt.get("operation_outcome") == outcome, "Local run footer differs from final native outcome")
    raw, record = receipt.get("run_manifest_raw"), receipt.get("run_manifest")
    need(isinstance(raw, str) and _unique_json(raw) == record and sha256(raw.encode("utf-8")).hexdigest() == receipt.get("run_manifest_sha256")
         == identities["WinBookSplit_Run.json"]["sha256"] and len(raw.encode("utf-8")) == identities["WinBookSplit_Run.json"]["size_bytes"],
         "Local manifest raw bytes/hash differ")
    need(record.get("protocol") == "winbooksplit.run" and type(record.get("version")) is int and record["version"] == 1
         and record.get("run_id") == owner["run_id"] and record.get("diagnostics_finalized") is True
         and record.get("application_version") == APP_VERSION and record.get("outcome") == outcome
         and record.get("engine_result") == outcome.get("engine_result"), "Local run protocol/owner/outcome differs")
    for key in ("started_utc", "finished_utc"):
        need(isinstance(record.get(key), str), "Local run timestamp missing")
    started, finished = (datetime.fromisoformat(record[key]) for key in ("started_utc", "finished_utc"))
    need(started.tzinfo is not None and finished.tzinfo is not None and started <= finished, "Local run timestamps invalid")
    versions, settings, logged = record.get("runtime_versions", {}), record.get("settings", {}), record.get("log", {})
    need(python_version_supported(versions.get("python")) and versions.get("pypdf") == RUNTIME_POLICY["pypdf"]["version"]
         and isinstance(versions.get("powershell"), str) and versions["powershell"].startswith(("5.1.", "7."))
         and (versions.get("calibre") is None or compatibility.calibre_version_supported(versions["calibre"], RUNTIME_POLICY)), "Local run runtime versions differ")
    if expected_python_version is not None:
        need(versions.get("python") == expected_python_version, "Local run Python differs from the selected probe")
    need(settings.get("mode") == outcome.get("mode") and settings.get("input_kind") in {"pdf", "epub", "azw3"}
         and all(type(settings.get(key)) is bool for key in ("preview", "non_interactive", "no_pause", "keep_converted_pdf"))
         and all(type(settings.get(key)) is int and settings[key] > 0 for key in ("conversion_timeout", "process_timeout")),
         "Local run normalized settings invalid")
    need(logged.get("filename") == "console.log" and type(logged.get("bytes_written")) is int
         and logged["bytes_written"] == receipt["log_size_bytes"] <= logged.get("max_bytes", -1) == 50331648
         and logged.get("body_limit_bytes") == 33554432 and type(logged.get("limit_reached")) is bool, "Local run bounded closed-log metrics differ")
    frame = outcome.get("engine_result") or {}
    plan = record.get("plan")
    expected_plan = frame.get("plan") or receipt.get("confirmed_plan")
    need(plan == expected_plan, "Local run plan differs from the actual prepared/confirmed plan")
    source = record.get("source_identity", {})
    if plan is not None:
        need(source == plan.get("original_ebook_identity", plan.get("source_identity"))
             and settings.get("normalized_inputs") == plan.get("normalized_inputs"), "Local run plan/source/settings differ")
    else:
        need(source.get("binding") == "metadata_only" and source.get("sha256") is None
             and settings.get("normalized_inputs") is None, "Local run invented an unavailable source/plan identity")
    if source_path is not None:
        need(source.get("path") == str(source_path) and Path(source_path).suffix.lower().removeprefix(".") == settings["input_kind"],
             "Local run source path/kind differs from actual input")
    if source_observation is not None:
        need(source.get("size_bytes") == source_observation.get("size_bytes")
             and (source.get("binding") == "metadata_only" or source.get("sha256") == source_observation.get("sha256")),
             "Local run source identity differs from actual immutable observation")
    parser = (frame.get("diagnostic") or {}).get("parser_warnings") or {}
    need(record.get("warnings") == {"planning": frame.get("warnings", []), "parser": parser.get("records", []),
         "parser_suppressed_count": parser.get("suppressed_count", 0), "parser_message_truncated_count": parser.get("message_truncated_count", 0)},
         "Local run warnings differ from engine categories")

def authenticate_console_manifest(log, outcome=None, *, source_path=None, source_observation=None, allow_unfinalized=False, expected_python_version=None):
    """Read only exact ordinary owned console records; never adopt pending files."""
    directory = log.parent
    require(log.is_absolute() and log.name == "console.log" and directory.resolve(strict=True) == directory
            and not history.is_reparse(directory) and re.fullmatch(r"\.WinBookSplit-console-[0-9a-f]{32}", directory.name),
            "Local console manifest containment/ownership path differs")
    directory_details = directory.lstat()
    names = {item.name for item in directory.iterdir()}
    absent = allow_unfinalized and names == CONSOLE_MEMBERS - {"WinBookSplit_Run.json"}
    require(names == CONSOLE_MEMBERS or absent, "Local console refuses unexpected or pending members")
    identities, raw = {}, {}
    for name in sorted(names):
        path = directory / name
        details = path.lstat()
        require(stat.S_ISREG(details.st_mode) and not history.is_reparse(path), "Local console member is not ordinary")
        data = path.read_bytes()
        raw[name] = data.decode("utf-8", errors="strict")
        identities[name] = {"device": details.st_dev, "inode": details.st_ino, "size_bytes": len(data), "sha256": sha256(data).hexdigest()}
    owner = _unique_json(raw[".WinBookSplit-console-owner.json"])
    require(owner == {"run_id": directory.name.removeprefix(".WinBookSplit-console-"), "kind": "console"}, "Local console marker differs")
    footers = [_unique_json(line[len("[OPERATION-OUTCOME] "):]) for line in raw["console.log"].splitlines() if line.startswith("[OPERATION-OUTCOME] ")]
    footer = footers[-1] if footers else None
    require(outcome is not None or footer is not None, "Local console lacks an authoritative operation outcome")
    outcome = footer if outcome is None else outcome
    manifest_raw = raw.get("WinBookSplit_Run.json")
    receipt = {"state": "unfinalized" if absent else "finalized", "members": sorted(names), "owner": owner,
               "directory_identity": {"device": directory_details.st_dev, "inode": directory_details.st_ino},
               "file_identities": identities, "log_sha256": identities["console.log"]["sha256"], "log_size_bytes": identities["console.log"]["size_bytes"],
               "operation_outcome": footer, "confirmed_plan": confirmed_console_plan(raw["console.log"]),
               "run_manifest": _unique_json(manifest_raw) if manifest_raw is not None else None, "run_manifest_raw": manifest_raw,
               "run_manifest_sha256": identities["WinBookSplit_Run.json"]["sha256"] if manifest_raw is not None else None}
    validate_console_evidence(receipt, outcome, receipt["log_sha256"], source_path=source_path,
                              source_observation=source_observation, allow_unfinalized=allow_unfinalized, log_text=raw["console.log"],
                              expected_python_version=expected_python_version)
    return receipt


def pdf_snapshot(path):
    from pypdf import PdfReader
    reader = PdfReader(ordinary(path))
    pages = []
    for page in reader.pages:
        stream = page.get_contents()
        geometry = {key: [float(value) for value in getattr(page, attr)] for key, attr in
                    (("media", "mediabox"), ("crop", "cropbox"), ("trim", "trimbox"), ("bleed", "bleedbox"), ("art", "artbox"))}
        geometry["rotation"] = page.rotation
        need(all(math.isfinite(value) for values in geometry.values() if isinstance(values, list) for value in values), "page_geometry_invalid")
        pages.append({"content_sha256": sha256(stream.get_data() if stream is not None else b"").hexdigest(),
                      "markers": re.findall(r"(?<![A-Z0-9-])WBS-PAGE-([0-9]{3})(?![0-9])", page.extract_text() or ""),
                      "geometry": geometry})
    return {"pages": pages, "metadata": dict(reader.metadata or {})}


def native_probe(probe):
    need(isinstance(probe, dict) and type(probe.get("Pid")) is int and probe["Pid"] > 0
         and type(probe.get("ExitCode")) is int, "native_process_missing")
    need(all(probe.get(key) is True for key in ("JobAssigned", "ParentStopped", "DescendantsStopped", "StreamsComplete", "InputWriterStopped")), "native_stop_unproved")
    need(all(probe.get(key) is False for key in ("TimedOut", "Cancelled", "StdoutTruncated", "StderrTruncated")), "native_transport_incomplete")
    need(all(key in probe and probe[key] is None for key in ("StartError", "StopError", "StreamError", "ResultError", "InteractionError", "InputError")), "native_transport_error")
    need(type(probe.get("Stdout")) is str and type(probe.get("Stderr")) is str, "native_streams_missing")
    need(type(probe.get("ElapsedSeconds")) in (int, float) and math.isfinite(probe["ElapsedSeconds"])
         and probe["ElapsedSeconds"] >= 0, "native_elapsed_invalid")
    for name in ("Stdout", "Stderr"):
        need(type(probe.get(name + "TotalBytes")) is int and probe[name + "TotalBytes"] == len(probe[name].encode("utf-8", errors="strict")), "native_stream_bytes_differ")
    return probe


def publication_once(directory, identities):
    details = ordinary(directory, directory=True).lstat()
    identity = (details.st_dev, details.st_ino)
    need(identity not in identities, "publication_directory_reused")
    identities.add(identity)


def outcome(stdout):
    lines = [line[len("[OUTCOME] "):] for line in stdout.splitlines() if line.startswith("[OUTCOME] ")]
    need(len(lines) == 1, "terminal_outcome_count")
    value = unique_json(lines[0])
    need(isinstance(value, dict) and set(value) == {"protocol", "version", "status", "code", "message", "exit_code", "mode", "written_count", "final_directory", "engine_result"}
         and value["protocol"] == "winbooksplit.outcome" and type(value["version"]) is int and value["version"] == 1
         and type(value["exit_code"]) is int and value["exit_code"] in {0, 2, 3, 4, 5, 6, 7, 130}
         and type(value["written_count"]) is int and value["written_count"] >= 0, "terminal_outcome_invalid")
    return value


def fixed_frame(stdout, prefix):
    lines = [line[len(prefix):] for line in stdout.splitlines() if line.startswith(prefix)]
    need(len(lines) == 1, "structured_frame_count")
    return unique_json(lines[0])


def validate_publication(base, execution, source, ranges, keep, version):
    """Adapt the existing exact-members, manifest and physical conversion oracle."""
    from pypdf import PdfReader
    need(isinstance(execution, dict) and re.fullmatch(r"[0-9a-f]{32}", execution.get("run_id", "")), "publication_run_invalid")
    final = ordinary(Path(execution["final_directory"]), directory=True)
    need(final.parent == base and re.search(r"_[0-9a-f]{32}$", final.name), "publication_escape")
    entries = execution.get("outputs")
    need(isinstance(entries, list) and len(entries) == execution.get("written_count") == len(ranges) > 0, "publication_count_invalid")
    need([[entry.get("start"), entry.get("end")] for entry in entries] == ranges, "publication_ranges_differ")
    need([number for start, end in ranges for number in range(start, end)] == list(range(len(source["pages"]))), "physical_coverage_invalid")
    names = [entry.get("filename") for entry in entries]
    need(all(type(name) is str and Path(name).name == name and name.endswith(".pdf") and ":" not in name for name in names)
         and len(set(name.casefold() for name in names)) == len(names), "publication_name_invalid")
    retained = execution.get("retained_intermediate")
    need((retained is not None) == keep, "converted_retention_differ")
    expected_members = {".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", *names}
    if keep:
        need(retained.get("filename") == "WinBookSplit_Converted.pdf", "converted_name_invalid")
        expected_members.add(retained["filename"])
    need({path.name for path in final.iterdir()} == expected_members, "publication_members_differ")
    for path in final.iterdir():
        ordinary(path)
    need(read_json(final / ".WinBookSplit-owner.json") == {"schema_version": 1, "kind": "run", "run_id": execution["run_id"]}, "publication_owner_differ")
    manifest = read_json(final / "WinBookSplit_Manifest.json")
    need(manifest == execution.get("manifest"), "publication_manifest_differ")
    keys = ("schema_version", "status", "run_id", "final_directory", "mode", "total_pages", "source_identity", "coverage", "written_count", "outputs")
    extra = {"original_ebook_identity", "conversion", "retained_intermediate"} if "original_ebook_identity" in execution else set()
    need(set(manifest) == set(keys) | extra, "publication_manifest_keys_differ")
    for key in keys:
        need(manifest.get(key) == (1 if key == "schema_version" else "complete" if key == "status" else execution.get(key)), "publication_manifest_field_differ")
    for key in ("original_ebook_identity", "conversion", "retained_intermediate"):
        need(manifest.get(key) == execution.get(key), "conversion_manifest_field_differ")
    for entry in entries:
        target = final / entry["filename"]
        actual = observation(target)
        reopened = pdf_snapshot(target)
        need(actual["sha256"] == entry.get("sha256") and actual["size_bytes"] == entry.get("size_bytes")
             and reopened["pages"] == source["pages"][entry["start"]:entry["end"]]
             and entry.get("page_count") == len(reopened["pages"]), "publication_physical_bytes_differ")
        metadata = reopened["metadata"]
        need(metadata.get("/Title") == entry.get("title") and metadata.get("/Creator") == "WinBookSplit"
             and metadata.get("/Producer") == "WinBookSplit " + version + " (pypdf " + RUNTIME_POLICY["pypdf"]["version"] + ")", "publication_metadata_differ")
        reader = PdfReader(target)
        need(len(reader.outline) == 1 and not isinstance(reader.outline[0], list)
             and reader.outline[0].title == entry["title"] and reader.get_destination_page_number(reader.outline[0]) == 0, "publication_bookmark_differ")
    if keep:
        target = final / retained["filename"]
        actual = observation(target)
        need(actual["sha256"] == retained.get("sha256") and actual["size_bytes"] == retained.get("size_bytes")
             and pdf_snapshot(target)["pages"] == source["pages"] and retained.get("page_count") == len(source["pages"]), "retained_pdf_differ")
    return final


def fixture_catalog(root):
    root = ordinary(root, directory=True)
    catalog = read_json(root / "native-fixtures.json")
    need(catalog.get("protocol") == "winbooksplit.native-fixtures" and type(catalog.get("schema_version")) is int and catalog["schema_version"] == 1
         and catalog.get("license") == "MIT" and catalog.get("authored_original") is True
         and catalog.get("remote_resources") is False, "fixture_provenance_invalid")
    expected = {"simple10", "nested12", "flat6", "epub", "azw3"}
    rows = catalog.get("files")
    need(isinstance(rows, dict) and set(rows) == expected, "fixture_members_differ")
    need(len({row.get("name", "").casefold() for row in rows.values()}) == len(rows), "fixture_filename_collision")
    for key, row in rows.items():
        need(isinstance(row, dict) and Path(row.get("name", "")).name == row.get("name") and row["name"], "fixture_name_invalid")
        actual = observation(root / row["name"])
        need(actual["sha256"] == row.get("sha256") and actual["size_bytes"] == row.get("size_bytes"), "fixture_hash_mismatch")
    need(compatibility.calibre_version_supported(catalog.get("calibre", {}).get("version"), RUNTIME_POLICY)
         and catalog.get("azw3_generation", {}).get("exit_code") == 0
         and catalog["azw3_generation"].get("job_assigned_before_resume") is True
         and catalog["azw3_generation"].get("descendants_stopped") is True, "genuine_azw3_missing")
    return catalog


def prepare(args):
    app = bind_application(args.application_root, args.policy_sha256)
    need(Path(sys.executable) == args.python_path and ordinary(args.python_path) == app / ".venv/Scripts/python.exe", "application_python_required")
    runtime = compatibility.probe_interpreter(policy_path=app / "engine/WinBookSplit.Compatibility.json")
    need(runtime.get("ok") is True and runtime.get("policy_sha256") == args.policy_sha256, "application_runtime_invalid")
    fixtures = fixture_catalog(args.fixture_root)
    work = ordinary(args.work_root, directory=True)
    need(not any(work.iterdir()), "case_work_not_empty")
    documents = ordinary(args.documents_root, directory=True)
    bare = ordinary(args.bare_python_path)
    calibre = ordinary(args.calibre_path)
    need(observation(calibre)["sha256"] == fixtures["calibre"].get("sha256"), "converter_fixture_identity_differ")
    sources = []
    for path in sorted(app.rglob("*")):
        if path.is_file() and not path.is_relative_to(app / ".venv"):
            sources.append(observation(path))
    cases, immutable = [], []
    for fmt in ("epub", "azw3"):
        reference = work / ("ref-" + fmt + ".pdf")
        source = args.fixture_root / fixtures["files"][fmt]["name"]
        reference_environment = {}
        settings = work / ("ref-" + fmt)
        settings.mkdir()
        for variable, name in (("CALIBRE_CONFIG_DIRECTORY", "cfg"), ("CALIBRE_CACHE_DIRECTORY", "cache"), ("TEMP", "tmp"), ("TMP", "tmp")):
            target = settings / name
            target.mkdir(exist_ok=True)
            reference_environment[variable] = str(target)
        cases.append({"id": "ref-" + fmt, "kind": "reference", "exe_selector": "calibre", "arguments": [str(source), str(reference), "--output-profile", "tablet"],
                      "cwd": str(work), "answers": [], "environment": reference_environment, "timeout_seconds": 240, "expected_exit": 0, "reference_path": str(reference), "format": fmt})
    manual_ranges = [[0, 3], [3, 6], [6, 10]]
    definitions = [("manual", "simple10", "Manual", "1,4,7", None, manual_ranges, False),
                   ("manual-first", "simple10", "Manual", "4,7", None, manual_ranges, False),
                   ("auto1", "nested12", "Auto", None, "1", [[0, 2], [2, 8], [8, 12]], False),
                   ("auto2", "nested12", "Auto", None, "2", [[0, 2], [2, 3], [3, 6], [6, 8], [8, 10], [10, 12]], False),
                   ("repeat", "simple10", "Manual", "1,4,7", None, manual_ranges, False),
                   ("preview", "simple10", "Manual", "1,4,7", None, manual_ranges, False),
                   ("invalid-manual", "simple10", "Manual", "7,04,4,1,abc", None, None, False),
                   ("no-bookmarks", "flat6", "Auto", None, "2", None, False),
                   ("cancel-plan", "simple10", "Manual", "1,4,7", None, manual_ranges, False),
                   ("missing-python", "simple10", "Manual", "1,4,7", None, None, False),
                   ("missing-pypdf", "simple10", "Manual", "1,4,7", None, None, False)]
    definitions += [(fmt + "-" + ("keep" if keep else "discard"), fmt, "Auto", None, "1", [[0, 1], [1, 2], [2, 3 if fmt == "epub" else 4]], keep)
                    for fmt in ("epub", "azw3") for keep in (False, True)]
    inputs = {}
    for index, (identifier, fixture, mode, starts, level, ranges, keep) in enumerate(definitions):
        if identifier == "repeat":
            source, base, cwd = inputs["manual"]
        else:
            folder = work / str(index)
            folder.mkdir()
            base, cwd = folder / "o", folder / "c"
            base.mkdir()
            cwd.mkdir()
            suffix = "." + (fixture if fixture in {"epub", "azw3"} else "pdf")
            source = folder / ("Å日 [1] & %WBS_KIT_X%!WBS_KIT_X!" + suffix)
            with source.open("xb") as stream:
                stream.write((args.fixture_root / fixtures["files"][fixture]["name"]).read_bytes())
            neighbor = source.with_suffix(".pdf") if suffix != ".pdf" else folder / "neighbor.txt"
            neighbor.write_bytes(b"Original synthetic source neighbor\n")
            prior = base / "prior-output.pdf"
            prior.write_bytes(b"Original synthetic prior output\n")
            for path in (source, neighbor, prior):
                immutable.append(observation(path))
            # Authored CWD decoys must never supply the isolated Python package.
            for name in ("pypdf.py", "winbooksplit_compatibility.py"):
                decoy = cwd / name
                decoy.write_bytes(b"raise RuntimeError('CWD_DECOY_MUST_NOT_EXECUTE')\n")
                immutable.append(observation(decoy))
        inputs[identifier] = (source, base, cwd)
        arguments = ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(app / "WinBookSplit.ps1"), "-InputFile", str(source), "-OutputDirectory", str(base),
                     "-PythonPath", str(work / "missing-python.exe") if identifier == "missing-python" else str(bare) if identifier == "missing-pypdf" else str(args.python_path), "-Mode", mode]
        arguments += ["-StartPages", starts] if starts else ["-BookmarkLevel", level]
        arguments += ["-NoPause"] if identifier == "cancel-plan" else ["-NonInteractive"]
        if identifier == "preview":
            arguments.append("-Preview")
        if fixture in {"epub", "azw3"}:
            arguments += ["-CalibrePath", str(calibre)]
        if keep:
            arguments.append("-KeepConvertedPdf")
        expected_exit = 2 if identifier == "invalid-manual" else NO_BOOKMARK_EXIT if identifier == "no-bookmarks" else 130 if identifier == "cancel-plan" else 3 if identifier.startswith("missing-") else 0
        cases.append({"id": identifier, "kind": "application", "exe_selector": "inbox_ps", "arguments": arguments, "cwd": str(cwd), "answers": ["C"] if identifier == "cancel-plan" else [],
                      "environment": {}, "timeout_seconds": 120, "expected_exit": expected_exit, "source_path": str(source), "output_base": str(base), "ranges": ranges,
                      "format": fixture if fixture in {"epub", "azw3"} else "pdf", "keep_converted": keep, "prior_members": ["prior-output.pdf"]})
    for count, identifier in ((1, "bat-one"), (2, "bat-two")):
        wrapper = work / (identifier + ".cmd")
        wrapper.write_bytes(("@echo off\r\nsetlocal DisableDelayedExpansion\r\n\"%WBS_KIT_BAT%\"" + "".join(' \"%WBS_KIT_INPUT' + str(index) + '%\"' for index in range(1, count + 1)) + "\r\n").encode("ascii"))
        immutable.append(observation(wrapper))
        source = inputs["manual"][0]
        cases.append({"id": identifier, "kind": "bat", "exe_selector": "cmd", "arguments": ["/d", "/v:off", "/c", str(wrapper)], "cwd": str(work),
                      "answers": ["1", "Y", "N", " "] if count == 1 else [" "], "environment": {"WBS_KIT_BAT": str(app / "WinBookSplit.bat"), "WBS_KIT_INPUT1": str(source),
                      "WBS_KIT_INPUT2": str(inputs["manual-first"][0]), "WBS_KIT_X": "MUST_NOT_EXPAND"}, "timeout_seconds": 180, "expected_exit": 0 if count == 1 else 2,
                      "source_path": str(source), "output_base": str(documents), "ranges": manual_ranges if count == 1 else None, "format": "pdf", "keep_converted": False})
    need(tuple(case["id"] for case in cases) == CASE_IDS, "planned_case_members_differ")
    return {"protocol": "winbooksplit.native-plan", "schema_version": 1, "application_version": APP_VERSION, "application_root": str(app), "fixture_root": str(args.fixture_root),
            "work_root": str(work), "python_path": str(args.python_path), "bare_python_path": str(bare), "calibre_path": str(calibre), "documents_root": str(documents),
            "policy_sha256": args.policy_sha256, "runtime": runtime, "calibre_version": fixtures["calibre"]["version"], "application_files": sources, "immutable_files": immutable,
            "fixture_manifest": observation(args.fixture_root / "native-fixtures.json"), "fixture_files": [observation(args.fixture_root / row["name"]) for row in fixtures["files"].values()], "cases": cases}


def validate_negative_outcome(identifier, final):
    expected_code = NEGATIVE_CODES.get(identifier)
    if expected_code:
        need(final["code"] == expected_code and final["written_count"] == 0 and final["final_directory"] is None,
             "negative_or_preview_outcome_differ")
    if identifier in {"no-bookmarks", "cancel-plan"}:
        frame = final.get("engine_result")
        expected_status = "no_plan" if identifier == "no-bookmarks" else "cancelled"
        need(final.get("status") == expected_status and isinstance(frame, dict)
             and frame.get("status") == expected_status and frame.get("code") == expected_code
             and "execution" in frame and frame["execution"] is None,
             "negative_engine_execution_differ")


def validate(plan, raw, plan_digest):
    need(plan.get("protocol") == "winbooksplit.native-plan" and type(plan.get("schema_version")) is int and plan["schema_version"] == 1
         and tuple(case.get("id") for case in plan.get("cases", [])) == CASE_IDS, "plan_schema_invalid")
    bind_application(Path(plan["application_root"]), plan["policy_sha256"])
    need(plan.get("application_version") == APP_VERSION, "application_version_changed")
    need(set(raw) == {"protocol", "schema_version", "plan_sha256", "cases"} and raw.get("protocol") == "winbooksplit.native-process"
         and type(raw.get("schema_version")) is int and raw["schema_version"] == 1 and raw.get("plan_sha256") == plan_digest, "process_report_binding_invalid")
    rows = raw.get("cases")
    need(isinstance(rows, list) and tuple(row.get("id") for row in rows) == CASE_IDS, "native_case_members_differ")
    for item in [*plan["application_files"], *plan["immutable_files"], plan["fixture_manifest"], *plan["fixture_files"]]:
        unchanged(item)
    runtime = plan["runtime"]
    need(compatibility.validate_interpreter(runtime, RUNTIME_POLICY)["ok"] and runtime.get("policy_sha256") == RUNTIME_POLICY_SHA256
         and runtime.get("pypdf_version") == RUNTIME_POLICY["pypdf"]["version"], "planned_runtime_invalid")
    need(compatibility.calibre_version_supported(plan.get("calibre_version"), RUNTIME_POLICY)
         and plan["calibre_version"] == fixture_catalog(Path(plan["fixture_root"]))["calibre"]["version"], "planned_converter_version_invalid")
    references, allowed_bases, results, publications = {}, {}, [], set()
    for case, row in zip(plan["cases"], rows):
        identifier = case["id"]
        try:
            need(row.get("status") == "COMPLETED", "native_case_not_completed")
            probe = native_probe(row.get("probe"))
            expected_exe = plan["calibre_path"] if case["exe_selector"] == "calibre" else str(Path(windows_system_directory()) / ("cmd.exe" if case["exe_selector"] == "cmd" else "WindowsPowerShell/v1.0/powershell.exe"))
            need(row.get("command") == [expected_exe, *case["arguments"]] and row.get("cwd") == case["cwd"] and row.get("answers") == case["answers"], "native_command_binding_invalid")
            need(probe["ExitCode"] == case["expected_exit"], "native_exit_differ")
            if case["environment"]:
                need(row.get("environment_before_after_restored") is True, "native_environment_not_restored")
            stdout = probe["Stdout"]
            if case["kind"] == "reference":
                reference = pdf_snapshot(Path(case["reference_path"]))
                need(len(reference["pages"]) == (3 if case["format"] == "epub" else 4)
                     and [marker for page in reference["pages"] for marker in page["markers"]] == ["001", "002", "003"], "conversion_reference_differ")
                from pypdf import PdfReader
                reader = PdfReader(Path(case["reference_path"]))
                need(len(reader.outline) == 3 and all(not isinstance(item, list) for item in reader.outline)
                     and [(item.title, reader.get_destination_page_number(item)) for item in reader.outline]
                     == [("Chapter One", 0), ("Chapter Two", 1), ("Chapter Three", 2)], "conversion_reference_bookmarks_differ")
                references[case["format"]] = reference
            elif identifier == "bat-two":
                need(row.get("environment_before_after_restored") is True and "[OUTCOME] " not in stdout
                     and "one" in stdout.lower() and "file" in stdout.lower(), "bat_refusal_differ")
            else:
                final = outcome(stdout)
                need(final["exit_code"] == probe["ExitCode"], "outcome_native_exit_differ")
                base = ordinary(Path(case["output_base"]), directory=True)
                if case["kind"] != "bat":
                    allowed_bases.setdefault(base, set(case["prior_members"]))
                frame = final.get("engine_result")
                validate_negative_outcome(identifier, final)
                if identifier.startswith("missing-"):
                    need(final["status"] == "failed" and final["written_count"] == 0 and final["final_directory"] is None
                         and final["code"] == NEGATIVE_CODES[identifier], "dependency_failure_differ")
                    dependency = fixed_frame(stdout, "[DEPENDENCY-ERROR] ")
                    need(dependency.get("Code") == final["code"] and dependency.get("Attempts"), "dependency_frame_missing")
                    if identifier == "missing-pypdf":
                        need(any(item.get("Reason") == "missing_pypdf" for item in dependency["Attempts"]), "genuine_missing_package_missing")
                elif final["exit_code"] == 0 and identifier != "preview":
                    need(final["status"] == "success" and final["code"] == "split_complete" and isinstance(frame, dict) and frame.get("status") == "success", "successful_outcome_differ")
                    source = pdf_snapshot(Path(case["source_path"])) if case["format"] == "pdf" else references[case["format"]]
                    ranges = case["ranges"]
                    published = validate_publication(base, frame["execution"], source, ranges, case["keep_converted"], APP_VERSION)
                    publication_once(published, publications)
                    need(final["written_count"] == len(ranges) and final["final_directory"] == str(published), "outcome_publication_differ")
                    if case["kind"] != "bat":
                        allowed_bases[base].add(published.name)
                    if case["format"] != "pdf":
                        generated = frame["execution"]["conversion"]["generated_pdf_identity"]["path"]
                        need(not Path(generated).exists(), "temporary_conversion_retained")
                if identifier == "preview":
                    need(frame.get("status") == "preview" and frame.get("execution") is None
                         and [[entry["start"], entry["end"]] for entry in frame["plan"]["entries"]] == case["ranges"], "preview_plan_differ")
                logs = [Path(line[len("Log: "):].strip()) for line in stdout.splitlines() if line.startswith("Log: ")]
                if not logs and identifier in {"cancel-plan", "no-bookmarks"}:
                    new_children = [path for path in base.iterdir() if path.name not in case["prior_members"]]
                    need(len(new_children) == 1 and re.fullmatch(r"\.WinBookSplit-console-[0-9a-f]{32}", new_children[0].name), "negative_console_ownership_missing")
                    logs = [new_children[0] / "console.log"]
                if final["exit_code"] == 0 and identifier != "preview" or identifier in {"cancel-plan", "no-bookmarks"}:
                    need(len(logs) == 1 and logs[0].parent.parent == base, "console_path_differ")
                    receipt = authenticate_console_manifest(logs[0], final, source_path=Path(case["source_path"]),
                                                            source_observation=observation(Path(case["source_path"])), expected_python_version=runtime["version"])
                    text = logs[0].read_text(encoding="utf-8")
                    dependency = fixed_frame(text, "[DEPENDENCY] ")["Runtime"]
                    need(dependency.get("Path") == plan["python_path"] and dependency.get("Details") == runtime
                         and dependency.get("InterpreterOnly") is False, "console_runtime_binding_differ")
                    native_probe(dependency.get("Probe"))
                    need(unique_json(dependency["Probe"]["Stdout"]) == runtime, "console_probe_binding_differ")
                    if case["kind"] != "bat":
                        allowed_bases[base].add(logs[0].parent.name)
                    if identifier in {"cancel-plan", "no-bookmarks"}:
                        need(receipt["confirmed_plan"] is None, "negative_consent_invented")
                else:
                    need(not logs, "unexpected_console_log")
                if case["kind"] == "bat":
                    need(row.get("environment_before_after_restored") is True, "bat_environment_not_restored")
            results.append({"id": identifier, "status": "PASS", "exit_code": probe["ExitCode"], "error_code": None})
        except (OracleError, ValueError, TypeError, KeyError, OSError) as error:
            # No arbitrary paths, messages, probe text or exception strings leak.
            results.append({"id": identifier, "status": "FAIL", "exit_code": row.get("probe", {}).get("ExitCode"),
                            "error_code": str(error) if isinstance(error, OracleError) and re.fullmatch(r"[a-z_]+", str(error)) else "oracle_input_invalid"})
    for base, allowed in allowed_bases.items():
        if {path.name for path in base.iterdir()} != allowed:
            for result in results:
                candidate = next(case for case in plan["cases"] if case["id"] == result["id"])
                if candidate.get("output_base") == str(base):
                    result.update(status="FAIL", error_code="output_base_members_differ")
    return {"protocol": "winbooksplit.native-oracle", "schema_version": 1, "success": all(row["status"] == "PASS" for row in results),
            "plan_sha256": plan_digest, "application_version": APP_VERSION, "policy_sha256": RUNTIME_POLICY_SHA256,
            "runtime_versions": {"python": runtime["version"], "pypdf": runtime["pypdf_version"], "calibre": plan["calibre_version"]},
            "source_unchanged": True, "cases": results, "limits": ["No Explorer gesture was performed.", "All native artifacts are preserved locally."]}


def export_report(raw):
    """Return only fixed fields; this function never copies arbitrary raw content."""
    oracle = raw.get("oracle")
    need(isinstance(oracle, dict) and oracle.get("protocol") == "winbooksplit.native-oracle", "redacted_oracle_missing")
    fields = {"protocol", "schema_version", "success", "plan_sha256", "application_version", "policy_sha256", "runtime_versions", "source_unchanged", "cases", "limits"}
    need(set(oracle) == fields and tuple(row.get("id") for row in oracle.get("cases", [])) == CASE_IDS, "redacted_oracle_schema_invalid")
    for row in oracle["cases"]:
        need(set(row) == {"id", "status", "exit_code", "error_code"} and row["status"] in {"PASS", "FAIL", "NOT_RUN", "BLOCKED"}
             and (row["exit_code"] is None or type(row["exit_code"]) is int)
             and (row["error_code"] is None or type(row["error_code"]) is str and re.fullmatch(r"[a-z_]{1,64}", row["error_code"])), "redacted_case_schema_invalid")
    need(all(type(oracle.get(key)) is bool for key in ("success", "source_unchanged"))
         and all(re.fullmatch(r"[0-9a-f]{64}", oracle.get(key, "")) for key in ("plan_sha256", "policy_sha256"))
         and type(oracle.get("schema_version")) is int and oracle["schema_version"] == 1
         and type(oracle.get("application_version")) is str and re.fullmatch(r"\d+\.\d+\.\d+", oracle["application_version"]), "redacted_identity_invalid")
    need(oracle["success"] == (oracle["source_unchanged"] and all(row["status"] == "PASS" for row in oracle["cases"])), "redacted_success_differ")
    need(set(oracle["runtime_versions"]) == {"python", "pypdf", "calibre"} and all(type(value) is str and re.fullmatch(r"\d+\.\d+\.\d+", value) for value in oracle["runtime_versions"].values()), "redacted_versions_invalid")
    need(oracle["limits"] == ["No Explorer gesture was performed.", "All native artifacts are preserved locally."], "redacted_limits_invalid")
    os_record = raw.get("os")
    need(isinstance(os_record, dict), "redacted_os_missing")
    safe_os = {key: os_record.get(key) for key in ("Caption", "Version", "BuildNumber", "OSArchitecture", "ProductType", "UBR", "powershell_version", "process_bits",
               "DisplayVersion", "processor_architectures", "os_bits", "platform_scope")}
    need(type(safe_os["Caption"]) is str and re.fullmatch(r"Microsoft Windows (?:10|11) (?:Home(?: Single Language| Country Specific| China)?|Pro(?: for Workstations| N for Workstations| Education)?|Enterprise(?: G)?|Education|IoT Enterprise|SE)(?: N)?", safe_os["Caption"])
         and type(safe_os["Version"]) is str and re.fullmatch(r"10\.0\.\d+", safe_os["Version"])
         and type(safe_os["BuildNumber"]) is str and re.fullmatch(r"\d{5}", safe_os["BuildNumber"])
         and safe_os["OSArchitecture"] in {"64-bit", "64ビット", "64 bits", "64-Bit", "64-bittinen"}
         and type(safe_os["ProductType"]) is int and safe_os["ProductType"] == 1
         and type(safe_os["UBR"]) is int and safe_os["UBR"] >= 0 and safe_os["process_bits"] == 64
         and type(safe_os["powershell_version"]) is str and re.fullmatch(r"5\.1\.\d+\.\d+", safe_os["powershell_version"]), "redacted_os_invalid")
    need(type(safe_os["process_bits"]) is int and type(safe_os["os_bits"]) is int and safe_os["os_bits"] == 64
         and type(safe_os["DisplayVersion"]) is str and re.fullmatch(r"\d{2}H[12]", safe_os["DisplayVersion"])
         and isinstance(safe_os["processor_architectures"], list) and 1 <= len(safe_os["processor_architectures"]) <= 64
         and all(type(value) is int and value == 9 for value in safe_os["processor_architectures"]), "redacted_native_architecture_invalid")
    scope = safe_os["platform_scope"]
    need(scope in {"NATIVE_WINDOWS10", "NATIVE_WINDOWS11_DRY_RUN", "BLOCKED"}, "redacted_platform_scope_invalid")
    need(safe_os["Version"] == "10.0." + safe_os["BuildNumber"], "redacted_os_build_incoherent")
    if scope == "NATIVE_WINDOWS10":
        need(safe_os["BuildNumber"] == "19045" and safe_os["Version"] == "10.0.19045" and safe_os["DisplayVersion"] == "22H2"
             and safe_os["Caption"].startswith("Microsoft Windows 10 "), "redacted_windows10_identity_invalid")
    elif scope == "NATIVE_WINDOWS11_DRY_RUN":
        need(int(safe_os["BuildNumber"]) >= 22000 and safe_os["Caption"].startswith("Microsoft Windows 11 "), "redacted_windows11_identity_invalid")
    else:
        need(oracle["success"] is False, "blocked_platform_cannot_pass")
    selection = raw.get("runtime_selection")
    selection_keys = {"owner_version", "owner_regular315_usable", "owner_selected", "selected_version",
                      "required_baseline_version", "baseline_version", "baseline_kind", "reason", "selected_identity", "baseline_identity"}
    need(isinstance(selection, dict) and set(selection) == selection_keys
         and type(selection["owner_regular315_usable"]) is bool and type(selection["owner_selected"]) is bool
         and selection["owner_selected"] == selection["owner_regular315_usable"]
         and (selection["owner_version"] is None or type(selection["owner_version"]) is str and re.fullmatch(r"3\.(?:13|14|15)\.\d+", selection["owner_version"]))
         and type(selection["selected_version"]) is str and re.fullmatch(r"3\.15\.\d+", selection["selected_version"])
         and selection["selected_version"] == oracle["runtime_versions"]["python"]
         and selection["required_baseline_version"] == selection["baseline_version"] == "3.15.0"
         and selection["baseline_kind"] in {"explicit3150_test_baseline", "existing3150_without_pypdf",
                                            "owned3150_venv_without_pip_from_existing3150", "reviewed_private3150_archive"}
         and selection["reason"] in {"usable_regular315_discovered", "supported_legacy_branch_is_not_315_kit_target", "no_usable_regular315_discovered"},
         "redacted_runtime_selection_invalid")
    identity_keys = {"source", "executable_sha256", "version", "implementation", "platform", "machine", "bits",
                     "gil_disabled", "releaselevel", "serial", "isolated", "dont_write_bytecode", "policy_sha256"}
    for name, version in (("selected_identity", selection["selected_version"]), ("baseline_identity", "3.15.0")):
        identity = selection[name]
        need(isinstance(identity, dict) and set(identity) == identity_keys
             and identity["source"] in {"explicit", "PATH", "registry", "manager", "py_launcher"}
             and identity["version"] == version and identity["implementation"] == "cpython"
             and identity["platform"] == "win32" and identity["machine"] == "AMD64"
             and type(identity["bits"]) is int and identity["bits"] == 64
             and identity["gil_disabled"] is False and identity["releaselevel"] == "final"
             and type(identity["serial"]) is int and identity["serial"] == 0
             and identity["isolated"] is True and identity["dont_write_bytecode"] is True
             and identity["policy_sha256"] == oracle["policy_sha256"]
             and type(identity["executable_sha256"]) is str and re.fullmatch(r"[0-9a-f]{64}", identity["executable_sha256"]),
             "redacted_runtime_identity_invalid")
    marked = raw.get("marked_setup_regressions")
    marked_keys = {"status", "required_runtime_version", "exact_asset_bound", "Action_roundtrips", "marked_BAT_routes",
                   "mocked_policy_refusals", "raw_receipt_sha256", "probe_sha256", "actual_inbox_policy_snapshots", "fixture", "routes"}
    need(isinstance(marked, dict) and set(marked) == marked_keys and marked["status"] == "PASS"
         and marked["required_runtime_version"] == "3.15.0" and marked["exact_asset_bound"] is True
         and all(type(marked[key]) is int for key in ("Action_roundtrips", "marked_BAT_routes", "mocked_policy_refusals"))
         and (marked["Action_roundtrips"], marked["marked_BAT_routes"], marked["mocked_policy_refusals"]) in {(2, 2, 1), (4, 3, 2)}
         and all(type(marked[key]) is str and re.fullmatch(r"[0-9a-f]{64}", marked[key]) for key in ("raw_receipt_sha256", "probe_sha256"))
         and marked["fixture"] == "Authored broken venv, not an owner legacy runtime; marks assigned by automation; all source/ADS/PDF checks verified",
         "redacted_marked_regressions_invalid")
    snapshots = marked["actual_inbox_policy_snapshots"]
    need(isinstance(snapshots, dict) and set(snapshots) == {"before", "after"} and snapshots["before"] == snapshots["after"], "redacted_policy_snapshots_invalid")
    scopes = snapshots["before"]
    need(isinstance(scopes, list) and len(scopes) == 5 and
         {row.get("scope") for row in scopes if isinstance(row, dict)} == {"MachinePolicy", "UserPolicy", "Process", "CurrentUser", "LocalMachine"}
         and all(set(row) == {"scope", "policy"} and row["policy"] in {"Undefined", "Restricted", "AllSigned", "RemoteSigned", "Unrestricted", "Bypass", "Default"} for row in scopes),
         "redacted_policy_snapshots_invalid")
    hosts = ("PS51", "PS766") if marked["Action_roundtrips"] == 4 else ("PS51",)
    expected_ids = {host + "-" + state for host in hosts for state in ("fresh", "stale")}
    expected_ids.update(host + "-marked-bat" for host in (*hosts, "CMD"))
    expected_ids.update(host + "-organization-simulation" for host in hosts)
    routes = marked["routes"]
    need(isinstance(routes, list) and len(routes) == len(expected_ids)
         and all(isinstance(row, dict) for row in routes) and {row.get("id") for row in routes} == expected_ids,
         "redacted_marked_routes_invalid")
    for route in routes:
        mock = route["id"].endswith("-organization-simulation")
        bat = route["id"].endswith("-marked-bat")
        keys = {"id", "status", "kind", "fixture", "exits", "child_hosts", "pdf"}
        if mock:
            keys.add("simulated_positive_launch_refused")
        need(set(route) == keys and route["status"] == "PASS"
             and route["kind"] == ("MOCKED_POLICY_QUERY" if mock else "BAT" if bat else "ACTION")
             and route["fixture"] == ("MOCKED_POLICY_QUERY" if mock else "AUTHORED_BROKEN_VENV" if bat else
                   "AUTHORED_BROKEN_VENV; no legacy Python execution" if route["id"].endswith("-stale") else "ACTUAL_BARE315"),
             "redacted_marked_route_schema_invalid")
        stages = ("simulation",) if mock else ("fresh", "check", "noop", "zeroargs-redirected", "broken-default", "repair", "pdf") if bat else ("direct", "before", "action", "retry")
        exits = route["exits"]
        need(isinstance(exits, list) and len(exits) == len(stages) and all(isinstance(row, dict) for row in exits)
             and tuple(row.get("stage") for row in exits) == stages
             and all(isinstance(row, dict) and set(row) == {"stage", "exit_code"} and type(row["exit_code"]) is int for row in exits),
             "redacted_marked_exit_schema_invalid")
        for row in exits:
            code = row["exit_code"]
            need(code == (1 if mock else 3 if row["stage"] in {"before", "broken-default"} else 0) if row["stage"] != "direct" else code != 0,
                 "redacted_marked_exit_differ")
        children = route["child_hosts"]
        child_stages = () if mock else stages[:-1] if bat else ("action",)
        need(isinstance(children, list) and len(children) == len(child_stages) and all(isinstance(row, dict) for row in children)
             and tuple(row.get("stage") for row in children) == child_stages,
             "redacted_child_host_schema_invalid")
        for child in children:
            need(isinstance(child, dict) and set(child) == {"stage", "executable_status", "version", "edition", "execution_policy", "selected_version", "pypdf_version"}
                 and child["executable_status"] == "observed" and child["edition"] == "Desktop"
                 and type(child["version"]) is str and re.fullmatch(r"5\.1\.\d+\.\d+", child["version"]), "redacted_child_host_schema_invalid")
            policy = child["execution_policy"]
            need(isinstance(policy, dict) and set(policy) == {"status", "effective", "scopes", "reason_code"}
                 and policy["status"] == "observed" and policy["effective"] == "Bypass" and policy["reason_code"] is None,
                 "redacted_child_policy_invalid")
            child_scopes = policy["scopes"]
            need(isinstance(child_scopes, list) and len(child_scopes) == 5
                 and all(isinstance(row, dict) and set(row) == {"scope", "policy"} for row in child_scopes)
                 and {row["scope"] for row in child_scopes} == {row["scope"] for row in scopes}, "redacted_child_policy_invalid")
            observed = {row["scope"]: row["policy"] for row in child_scopes}
            expected = {row["scope"]: row["policy"] for row in scopes}
            expected["Process"] = "Bypass"
            need(observed == expected, "redacted_child_policy_changed")
            broken = child["stage"] == "broken-default"
            # Historical export checks report coherence without binding a live policy.
            need(child["selected_version"] == (None if broken else "3.15.0")
                 and child["pypdf_version"] == (None if broken else oracle["runtime_versions"]["pypdf"]), "redacted_child_runtime_invalid")
        if mock:
            need(route["simulated_positive_launch_refused"] is True and route["pdf"] is None, "redacted_mock_refusal_invalid")
        else:
            pdf = route["pdf"]
            need(isinstance(pdf, dict) and set(pdf) == {"success", "runtime_version", "written_count", "physical_pages", "source_sha256"}
                 and pdf["success"] is True and pdf["runtime_version"] == "3.15.0"
                 and type(pdf["written_count"]) is int and pdf["written_count"] == 3
                 and type(pdf["physical_pages"]) is int and pdf["physical_pages"] == 10
                 and type(pdf["source_sha256"]) is str and re.fullmatch(r"[0-9a-f]{64}", pdf["source_sha256"]),
                 "redacted_marked_pdf_invalid")
    return {"protocol": "winbooksplit.native-result", "schema_version": 1, "evidence_kind": "NATIVE_AUTOMATED", "execution": raw.get("execution") if raw.get("execution") in {"AGENT_LAUNCHED", "USER_LAUNCHED_AUTOMATED"} else "UNKNOWN",
            "os": safe_os, "oracle": oracle, "runtime_selection": selection, "marked_setup_regressions": marked}


def prepare_store(args):
    """Finite PDF-only automatic recipes; inventory is never an application input."""
    app = bind_application(args.application_root, args.policy_sha256)
    python = ordinary(args.python_path)
    need(Path(sys.executable) == python == app / ".venv/Scripts/python.exe", "application_python_required")
    runtime = compatibility.probe_interpreter(policy_path=app / "engine/WinBookSplit.Compatibility.json")
    need(runtime.get("ok") is True, "application_runtime_invalid")
    fixtures = fixture_catalog(args.fixture_root)
    work = ordinary(args.work_root, directory=True)
    need(not any(work.iterdir()), "case_work_not_empty")
    documents = ordinary(args.documents_root, directory=True)
    cases, immutable = [], []
    definitions = (("manual", "simple10", "Manual", "1,4,7", None, [[0, 3], [3, 6], [6, 10]]),
                   ("auto1", "nested12", "Auto", None, "1", [[0, 2], [2, 8], [8, 12]]),
                   ("auto2", "nested12", "Auto", None, "2", [[0, 2], [2, 3], [3, 6], [6, 8], [8, 10], [10, 12]]),
                   ("bat", "simple10", None, None, None, [[0, 3], [3, 6], [6, 10]]))
    for identifier, fixture, mode, starts, level, ranges in definitions:
        folder = work / identifier
        folder.mkdir()
        cwd, base, source = folder / "unrelated-cwd", folder / "output", folder / "Authored Å日.pdf"
        cwd.mkdir()
        base.mkdir()
        source.write_bytes((args.fixture_root / fixtures["files"][fixture]["name"]).read_bytes())
        neighbor, prior = folder / "neighbor.txt", base / "prior-output.pdf"
        neighbor.write_bytes(b"Authored neighbor must remain unchanged\n")
        prior.write_bytes(b"Authored prior output must remain unchanged\n")
        for path in (source, neighbor, prior):
            immutable.append(observation(path))
        for name in ("pypdf.py", "winbooksplit_compatibility.py"):
            decoy = cwd / name
            decoy.write_bytes(b"raise RuntimeError('STORE_CWD_DECOY_MUST_NOT_EXECUTE')\n")
            immutable.append(observation(decoy))
        if identifier == "bat":
            # No controller environment is supplied, including benign BAT path
            # variables. Refuse cmd expansion in an unusual owner path instead.
            need(not any(character in str(app) + str(source) for character in '%\r\n"'), "store_bat_path_requires_literal_route")
            wrapper = folder / "ordinary-bat.cmd"
            # cmd reads batch files in its current code page. ASCII-only wrapper
            # paths preserve Unicode arguments through the native wide argv.
            wrapper.write_bytes(b'@echo off\r\nsetlocal DisableDelayedExpansion\r\nchcp 65001 >nul\r\n' + ('"' + str(app / "WinBookSplit.bat") + '" "' + str(source) + '"\r\n').encode("utf-8"))
            immutable.append(observation(wrapper))
            arguments, selector, answers, base = ["/d", "/v:off", "/c", str(wrapper)], "cmd", ["1", "Y", "N", " "], documents
        else:
            arguments = ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(app / "WinBookSplit.ps1"),
                         "-InputFile", str(source), "-OutputDirectory", str(base), "-Mode", mode]
            arguments += ["-StartPages", starts] if starts else ["-BookmarkLevel", level]
            arguments += ["-NonInteractive"]
            selector, answers = "inbox_ps", []
        need("-PythonPath" not in arguments, "store_runtime_override_forbidden")
        cases.append({"id": identifier, "exe_selector": selector, "arguments": arguments, "cwd": str(cwd), "answers": answers,
                      "environment": {}, "timeout_seconds": 180, "source_path": str(source), "output_base": str(base),
                      "ranges": ranges, "prior_members": ["prior-output.pdf"] if identifier != "bat" else []})
    return {"protocol": "winbooksplit.native-store-plan", "schema_version": 2, "application_root": str(app),
            "python_path": str(python), "runtime": runtime, "policy_sha256": args.policy_sha256,
            "immutable_files": immutable, "cases": cases}


def validate_store(plan, raw, plan_digest):
    need(plan.get("protocol") == "winbooksplit.native-store-plan" and type(plan.get("schema_version")) is int
         and plan["schema_version"] == 2, "store_plan_schema_invalid")
    need(tuple(case.get("id") for case in plan.get("cases", [])) == ("manual", "auto1", "auto2", "bat"), "store_plan_members_invalid")
    need(raw.get("protocol") == "winbooksplit.native-store-process" and raw.get("schema_version") == 2
         and type(raw.get("schema_version")) is int and raw.get("plan_sha256") == plan_digest
         and len(raw.get("cases", [])) == 4, "store_process_binding_invalid")
    app = bind_application(Path(plan["application_root"]), plan["policy_sha256"])
    need(plan["python_path"] == str(app / ".venv/Scripts/python.exe"), "store_venv_identity_invalid")
    for item in plan["immutable_files"]:
        unchanged(item)
    results, publications = [], set()
    for case, row in zip(plan["cases"], raw["cases"]):
        expected_exe = Path(windows_system_directory()) / ("cmd.exe" if case["exe_selector"] == "cmd" else "WindowsPowerShell/v1.0/powershell.exe")
        need(row.get("id") == case["id"] and row.get("command") == [str(expected_exe), *case["arguments"]]
             and row.get("cwd") == case["cwd"] and row.get("answers") == case["answers"]
             and row.get("environment_before_after_restored") is True and case["environment"] == {}
             and row.get("failure") is None and "-PythonPath" not in case["arguments"], "store_process_recipe_differ")
        probe = row["probe"]
        native_probe(probe)
        need(probe["ExitCode"] == 0, "store_application_failed")
        final = outcome(probe["Stdout"])
        need(final["status"] == "success" and final["code"] == "split_complete" and final["exit_code"] == 0, "store_outcome_failed")
        source = Path(case["source_path"])
        base = Path(case["output_base"])
        published = validate_publication(base, final["engine_result"]["execution"], pdf_snapshot(source), case["ranges"], False, APP_VERSION)
        publication_once(published, publications)
        need(final["written_count"] == len(case["ranges"]) and final["final_directory"] == str(published), "store_output_differ")
        logs = [Path(line[5:].strip()) for line in probe["Stdout"].splitlines() if line.startswith("Log: ")]
        need(len(logs) == 1 and logs[0].parent.parent == base, "store_console_path_differ")
        receipt = authenticate_console_manifest(logs[0], final, source_path=source, source_observation=observation(source),
                                                expected_python_version=plan["runtime"]["version"])
        dependency = fixed_frame(logs[0].read_text(encoding="utf-8"), "[DEPENDENCY] ")["Runtime"]
        need(dependency.get("Path") == plan["python_path"] and dependency.get("Details") == plan["runtime"]
             and dependency.get("Source") == "application_venv", "store_console_runtime_differ")
        if case["id"] != "bat":
            need({child.name for child in base.iterdir()} == set(case["prior_members"]) | {published.name, logs[0].parent.name}, "store_output_members_differ")
        results.append({"id": case["id"], "status": "PASS", "exit_code": 0, "written_count": final["written_count"],
                        "full_page_coverage": True, "console_receipt": receipt})
    return {"protocol": "winbooksplit.native-store-oracle", "schema_version": 2, "success": True,
            "source_unchanged": True, "plan_sha256": plan_digest, "cases": results}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    p = sub.add_parser("prepare")
    for name in ("application-root", "fixture-root", "work-root", "python-path", "bare-python-path", "calibre-path", "documents-root"):
        p.add_argument("--" + name, type=Path, required=True)
    p.add_argument("--policy-sha256", required=True)
    v = sub.add_parser("validate")
    v.add_argument("--plan", type=Path, required=True)
    v.add_argument("--process-report", type=Path, required=True)
    e = sub.add_parser("export")
    e.add_argument("--raw-report", type=Path, required=True)
    sp = sub.add_parser("prepare-store")
    for name in ("application-root", "fixture-root", "work-root", "python-path", "documents-root"):
        sp.add_argument("--" + name, type=Path, required=True)
    sp.add_argument("--policy-sha256", required=True)
    sv = sub.add_parser("validate-store")
    sv.add_argument("--plan", type=Path, required=True)
    sv.add_argument("--process-report", type=Path, required=True)
    for child in (p, v, e, sp, sv):
        child.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        need(sys.flags.isolated and sys.dont_write_bytecode, "isolated_python_required")
        if args.command == "prepare":
            value = prepare(args)
        elif args.command == "prepare-store":
            value = prepare_store(args)
        elif args.command == "validate-store":
            value = validate_store(read_json(args.plan), read_json(args.process_report), observation(args.plan)["sha256"])
        elif args.command == "validate":
            value = validate(read_json(args.plan), read_json(args.process_report), observation(args.plan)["sha256"])
        else:
            value = export_report(read_json(args.raw_report))
        write_json(args.output, value)
        return 0 if value.get("success", value.get("oracle", {}).get("success", True)) else 1
    except (OracleError, OSError, ValueError, TypeError, KeyError) as error:
        code = str(error) if isinstance(error, OracleError) and re.fullmatch(r"[a-z_]+", str(error)) else "oracle_input_invalid"
        print(json.dumps({"protocol": "winbooksplit.native-oracle-error", "schema_version": 1, "success": False, "code": code}))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
