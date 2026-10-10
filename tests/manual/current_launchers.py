"""Current transactional launchers using the unchanged historical process guards."""

from concurrent.futures import ThreadPoolExecutor
import json
import os
from pathlib import Path
import re
import shutil
import stat
import sys
import uuid
import importlib.util
from datetime import datetime
from hashlib import sha256

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("wbs_current_launcher_manual", ROOT / "tests/manual/characterize_manual.py")
manual = importlib.util.module_from_spec(spec)
spec.loader.exec_module(manual)
history, require = manual.history, manual.require
spec = importlib.util.spec_from_file_location("wbs_current_interaction_receipts", ROOT / "tests/manual/interaction_receipts.py")
interactions = importlib.util.module_from_spec(spec)
spec.loader.exec_module(interactions)


def exact_line(stdout, prefix):
    values = [line[len(prefix):].strip() for line in stdout.splitlines() if line.startswith(prefix)]
    require(len(values) == 1 and values[0], "Exactly one explicit launcher " + prefix + " path required")
    return Path(values[0])


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


def validate_console_evidence(receipt, outcome, log_sha256, *, source_path=None, source_observation=None, allow_unfinalized=False):
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
         and record.get("application_version") == "1.0.0-dev" and record.get("outcome") == outcome
         and record.get("engine_result") == outcome.get("engine_result"), "Local run protocol/owner/outcome differs")
    for key in ("started_utc", "finished_utc"):
        need(isinstance(record.get(key), str), "Local run timestamp missing")
    started, finished = (datetime.fromisoformat(record[key]) for key in ("started_utc", "finished_utc"))
    need(started.tzinfo is not None and finished.tzinfo is not None and started <= finished, "Local run timestamps invalid")
    versions, settings, logged = record.get("runtime_versions", {}), record.get("settings", {}), record.get("log", {})
    need(versions.get("python") == "3.14.8" and versions.get("pypdf") == "6.19.0"
         and isinstance(versions.get("powershell"), str) and versions["powershell"].startswith(("5.1.", "7."))
         and versions.get("calibre") in {None, "9.15.0"}, "Local run runtime versions differ")
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


def authenticate_console_manifest(log, outcome=None, *, source_path=None, source_observation=None, allow_unfinalized=False):
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
    plans = [_unique_json(line[len("[PLAN] "):])["plan"] for line in raw["console.log"].splitlines() if line.startswith("[PLAN] ")]
    footer = footers[-1] if footers else None
    require(outcome is not None or footer is not None, "Local console lacks an authoritative operation outcome")
    outcome = footer if outcome is None else outcome
    manifest_raw = raw.get("WinBookSplit_Run.json")
    receipt = {"state": "unfinalized" if absent else "finalized", "members": sorted(names), "owner": owner,
               "directory_identity": {"device": directory_details.st_dev, "inode": directory_details.st_ino},
               "file_identities": identities, "log_sha256": identities["console.log"]["sha256"], "log_size_bytes": identities["console.log"]["size_bytes"],
               "operation_outcome": footer, "confirmed_plan": plans[-1] if plans else None,
               "run_manifest": _unique_json(manifest_raw) if manifest_raw is not None else None, "run_manifest_raw": manifest_raw,
               "run_manifest_sha256": identities["WinBookSplit_Run.json"]["sha256"] if manifest_raw is not None else None}
    validate_console_evidence(receipt, outcome, receipt["log_sha256"], source_path=source_path,
                              source_observation=source_observation, allow_unfinalized=allow_unfinalized)
    return receipt


def remove_console_directory(log, parent, outcome=None, **observations):
    receipt = authenticate_console_manifest(log, outcome, **observations)
    remove_known_directory(log.parent, parent, set(receipt["members"]), ".WinBookSplit-console-owner.json", receipt["owner"],
                           expected_sha256={name: item["sha256"] for name, item in receipt["file_identities"].items()},
                           expected_file_identities=receipt["file_identities"], expected_directory_identity=receipt["directory_identity"])
    return receipt


def remove_known_directory(directory, parent, members, marker_name, marker, *, expected_sha256=None,
                           expected_file_identities=None, expected_directory_identity=None):
    """Delete held matching objects; a path check never authorizes unlink."""
    require(directory.is_absolute() and directory.parent == parent and directory.resolve(strict=True) == directory
            and not history.is_reparse(directory), "Current cleanup containment/reparse guard failed")
    details = directory.lstat()
    directory_identity = details.st_dev, details.st_ino
    if expected_directory_identity is not None:
        require(directory_identity == (expected_directory_identity["device"], expected_directory_identity["inode"]),
                "Current cleanup authenticated directory identity changed before held deletion")
    if expected_file_identities is not None:
        require(set(expected_file_identities) == set(members), "Current cleanup authenticated member identity scope differs")
    actual = list(directory.iterdir())
    require({path.name for path in actual} == set(members), "Current cleanup refuses unexpected members")
    identities = {}
    for path in actual:
        details = path.lstat()
        require(not details.st_file_attributes & stat.FILE_ATTRIBUTE_REPARSE_POINT and stat.S_ISREG(details.st_mode),
                "Current cleanup refuses reparse or nonregular members")
        identities[path] = details.st_dev, details.st_ino
        if expected_file_identities is not None:
            expected = expected_file_identities[path.name]
            require(identities[path] == (expected["device"], expected["inode"]),
                    "Current cleanup authenticated member identity changed before held deletion")
    require(json.loads((directory / marker_name).read_text(encoding="utf-8")) == marker,
            "Current cleanup ownership marker changed")
    windows = manual.load_module("wbs_current_launcher_windows", ROOT / "engine/winbooksplit_windows.py")
    directories, read_files, delete_files = [], [], []
    try:
        # Retain root-to-parent ancestors before holding the exact owned child.
        for ancestor in [*reversed(parent.parents), parent]:
            directories.append(windows.DirectoryGuard(ancestor))
        owned = windows.DirectoryGuard(directory, delete_access=True)
        directories.append(owned)
        require(owned.stat_identity == directory_identity and directory.resolve(strict=True) == directory,
                "Current cleanup refuses a replaced directory")
        for path in actual:
            guard = windows.FileGuard(path)
            read_files.append(guard)
            require(guard.stat_identity == identities[path], "Current cleanup refuses a replaced member")
        require({path.name for path in directory.iterdir()} == set(members)
                and json.loads((directory / marker_name).read_text(encoding="utf-8")) == marker,
                "Current cleanup membership or ownership changed while acquiring read locks")
        if expected_sha256 is not None:
            require(set(expected_sha256) == set(members) and all(sha256(path.read_bytes()).hexdigest() == expected_sha256[path.name] for path in actual),
                    "Current cleanup authenticated console bytes changed before held deletion")
        held_identities = {path: guard.identity for path, guard in zip(actual, read_files)}
        for guard in directories + read_files:
            guard.assert_unchanged()
        # CRT byte reads do not share DELETE on Windows. Validate bytes under
        # read-only locks, then reacquire matching delete handles without reads.
        for guard in read_files:
            guard.close()
        read_files.clear()
        for path in actual:
            guard = windows.FileGuard(path, delete_access=True)
            delete_files.append(guard)
            require(guard.stat_identity == identities[path] and guard.identity == held_identities[path],
                    "Current cleanup refuses a member replaced while acquiring delete handles")
        require({path.name for path in directory.iterdir()} == set(members),
                "Current cleanup membership changed while acquiring delete handles")
        for guard in directories + delete_files:
            guard.assert_unchanged()
        for guard in delete_files:
            guard.delete()
        owned.delete()  # Windows rejects a new foreign member/nonempty child.
        require(not os.path.lexists(directory), "Owned current output directory remained after held deletion")
    finally:
        for guard in reversed(delete_files + read_files):
            guard.close()
        for guard in reversed(directories):
            guard.close()


def launchers(work, fixtures, generator, shells, cases):
    require(os.name == "nt" and len(shells) == 2, "Actual PS5.1 and PS7 required")
    system_root = Path(os.environ["SystemRoot"])
    ps51 = system_root / "System32/WindowsPowerShell/v1.0/powershell.exe"
    require(ps51.resolve() in [path.resolve() for path in shells]
            and all(path.is_absolute() and path.is_file() for path in shells), "Explicit supported hosts required")
    probe = history.run([str(ps51), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-Command",
                         "[Environment]::GetFolderPath('MyDocuments')"], work,
                        environment={**history.clean_environment(work), "PSModulePath": str(ps51.parent / "Modules")})
    require(probe["exit_code"] == 0 and probe["stdout"].strip(), "Documents observation failed")
    documents = Path(probe["stdout"].strip()).resolve(strict=True)
    require(documents.is_dir() and not history.is_reparse(documents), "Actual Documents must be a normal directory")
    shared = work / "shared-temp"
    shared.mkdir()
    sentinel = shared / "WinBookSplit_Engine_v2.py"
    sentinel.write_bytes(b"Original owned shared-temp engine-name sentinel\n")
    sentinel_hash = history.file_digest(sentinel)
    references = {case["oracle_id"]: case["original"]["outputs"] for case in cases}

    def one(label, host, fixture_name, oracle_id, batch=False):
        token = uuid.uuid4().hex
        cwd = work / label
        cwd.mkdir()
        counterfeit = cwd / "engine"
        counterfeit.mkdir()
        decoy = counterfeit / "winbooksplit_engine.py"
        decoy.write_text("raise RuntimeError('untrusted CWD engine executed')\n", encoding="utf-8")
        decoy_hash = history.file_digest(decoy)
        source = cwd / ("wbs-m1-" + token + ".pdf")
        shutil.copyfile(fixtures[fixture_name], source)
        source_hash = history.file_digest(source)
        environment = history.clean_environment(shared)
        environment["PATH"] = os.pathsep.join((str(Path(sys.executable).parent), str(ps51.parent),
                                               str(system_root / "System32"), str(system_root)))
        environment["PSModulePath"] = str(host.parent / "Modules")
        stdin = "M\n4,7\nY\nN\n\n" if oracle_id == "MAN-03" else "2\nY\nN\n\n"
        if batch:
            require(not any(character in str(ROOT / "WinBookSplit.bat") + str(source)
                            for character in ' %!&|<>^"'), "Generated batch paths must remain safe ASCII")
            argv = [str(system_root / "System32/cmd.exe"), "/d", "/c", str(ROOT / "WinBookSplit.bat"), str(source)]
        else:
            argv = [str(host), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", str(ROOT / "WinBookSplit.ps1"), str(source)]
        parent_sentinel = history.owned_document_directory(documents, token)
        cleanup_safe, final, log, log_validated = True, None, None, False
        members = None
        try:
            try:
                result = history.run_entrypoint(argv, cwd, environment=environment, stdin=stdin)
            except history.EntryPointFailure as error:
                cleanup_safe = error.cleanup_safe
                if not cleanup_safe:
                    raise RuntimeError(str(error) + "; preserve owned marked launcher paths; PID " + str(error.pid)) from error
                raise
            require(result["exit_code"] == 0, "Actual launcher failed: " + label + ": " + repr(result))
            final, log = exact_line(result["stdout"], "Output: "), exact_line(result["stdout"], "Log: ")
            require(final.parent == documents and final.name.startswith(source.stem + "_")
                    and final.resolve(strict=True) == final and not history.is_reparse(final), "Launcher final path lost owned input prefix")
            manifest = json.loads((final / "WinBookSplit_Manifest.json").read_text(encoding="utf-8"))
            execution = {key: manifest[key] for key in ("run_id", "final_directory", "mode", "total_pages", "source_identity", "coverage", "written_count", "outputs")}
            execution.update(manifest_filename="WinBookSplit_Manifest.json", manifest=manifest)
            actual = manual.published_outputs(documents, generator, execution)
            require(actual == references[oracle_id], "Actual launcher bytes/page identities differ from corrected reference")
            members = {".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", *(item["filename"] for item in actual)}
            require(log.is_absolute() and log.name == "console.log" and log.parent.parent == documents
                    and re.fullmatch(r"\.WinBookSplit-console-[0-9a-f]{32}", log.parent.name)
                    and log.parent.resolve(strict=True) == log.parent and not history.is_reparse(log.parent), "Console path ownership failed")
            owner = json.loads((log.parent / ".WinBookSplit-console-owner.json").read_text(encoding="utf-8"))
            require(owner == {"run_id": log.parent.name.removeprefix(".WinBookSplit-console-"), "kind": "console"}, "Console marker differs")
            log_validated = True
            log_text = log.read_text(encoding="utf-8")
            console_evidence = authenticate_console_manifest(log, source_path=source,
                source_observation={"sha256": source_hash, "size_bytes": source.stat().st_size})
            interaction = interactions.capture(log_text)
            summaries = [json.loads(line[len("[PROCESS] "):]) for line in log_text.splitlines() if line.startswith("[PROCESS] ")]
            interactions.validate(interaction, summaries,
                ["input_ready", "plan_ready"] if oracle_id == "MAN-03" else ["plan_ready"],
                ["starts", "execute"] if oracle_id == "MAN-03" else ["execute"], starts="4,7", execution=execution)
            require(source_hash == history.file_digest(source) and decoy_hash == history.file_digest(decoy)
                    and (parent_sentinel / ".wbs-m1-neighbor.txt").read_bytes() == b"Original synthetic Documents neighbor\n", "Launcher modified input/neighbor")
            version = history.run([str(host), "-NoProfile", "-Command", "$PSVersionTable.PSVersion.ToString()"], cwd, environment=environment)
            require(version["exit_code"] == 0, "Host observation failed")
            return {"id": label, "entrypoint": "BAT -> actual Windows PowerShell" if batch else "PowerShell -File",
                    "host": version["stdout"].strip(), "shell_executable": str(host), "controlled_python": sys.executable,
                    "builtin_module_path": environment["PSModulePath"], "python_path_first": str(Path(sys.executable).parent),
                    "oracle_id": oracle_id, "outputs": actual, "input_unchanged": True, "cwd_engine_untouched": True,
                    "owned_neighbor_unchanged": True, "run_id": execution["run_id"], "manifest_validated": True,
                    "stdin_utf8": stdin, "interaction": interaction, "process_summaries": summaries,
                    "console_evidence": console_evidence,
                    "output_location": "<observed-Documents>/<owned-input-GUID>_<timestamp>_<run-GUID>",
                    **{key: value.replace(str(documents), "<observed-Documents>") if isinstance(value, str) else value
                       for key, value in result.items()}}
        finally:
            if cleanup_safe:
                if final is not None and members is not None:
                    remove_known_directory(final, documents, members, ".WinBookSplit-owner.json",
                                           {"schema_version": 1, "kind": "run", "run_id": execution["run_id"]})
                if log is not None and log_validated:
                    remove_console_directory(log, documents, source_path=source,
                        source_observation={"sha256": source_hash, "size_bytes": source.stat().st_size})
                history.cleanup_document_directory(parent_sentinel, documents, token, set())

    specifications = [("PS51-unrelated", ps51, "simple10", "MAN-03", False),
                      ("PS7-unrelated", next(path for path in shells if path.resolve() != ps51.resolve()), "nested12", "BM-03", False),
                      ("BAT-unrelated", ps51, "simple10", "MAN-03", True)]
    records = [one(*item) for item in specifications]
    with ThreadPoolExecutor(max_workers=3) as executor:
        futures = [executor.submit(one, "parallel-" + item[0], *item[1:]) for item in specifications]
        records.extend(future.result() for future in futures)
    require(history.file_digest(sentinel) == sentinel_hash, "Shared-temp sentinel changed")
    return {"probes": records, "probe_count": 6, "parallel_launch_count": 3,
            "shared_temp_engine_sentinel_unchanged": True, "owned_document_outputs_removed": True,
            "method": "Actual PS5.1/PS7/BAT, controlled Python PATH, exact emitted manifest/log paths, unchanged historical process and GUID-parent guards; only validated known direct children removed."}
