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

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("wbs_current_launcher_manual", ROOT / "tests/manual/characterize_manual.py")
manual = importlib.util.module_from_spec(spec)
spec.loader.exec_module(manual)
history, require = manual.history, manual.require


def exact_line(stdout, prefix):
    values = [line[len(prefix):].strip() for line in stdout.splitlines() if line.startswith(prefix)]
    require(len(values) == 1 and values[0], "Exactly one explicit launcher " + prefix + " path required")
    return Path(values[0])


def remove_known_directory(directory, parent, members, marker_name, marker):
    """Delete held matching objects; a path check never authorizes unlink."""
    require(directory.is_absolute() and directory.parent == parent and directory.resolve(strict=True) == directory
            and not history.is_reparse(directory), "Current cleanup containment/reparse guard failed")
    details = directory.lstat()
    directory_identity = details.st_dev, details.st_ino
    actual = list(directory.iterdir())
    require({path.name for path in actual} == set(members), "Current cleanup refuses unexpected members")
    identities = {}
    for path in actual:
        details = path.lstat()
        require(not details.st_file_attributes & stat.FILE_ATTRIBUTE_REPARSE_POINT and stat.S_ISREG(details.st_mode),
                "Current cleanup refuses reparse or nonregular members")
        identities[path] = details.st_dev, details.st_ino
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
        stdin = "M\n4,7\n\n" if oracle_id == "MAN-03" else "2\n\n"
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
            require(source_hash == history.file_digest(source) and decoy_hash == history.file_digest(decoy)
                    and (parent_sentinel / ".wbs-m1-neighbor.txt").read_bytes() == b"Original synthetic Documents neighbor\n", "Launcher modified input/neighbor")
            version = history.run([str(host), "-NoProfile", "-Command", "$PSVersionTable.PSVersion.ToString()"], cwd, environment=environment)
            require(version["exit_code"] == 0, "Host observation failed")
            return {"id": label, "entrypoint": "BAT -> actual Windows PowerShell" if batch else "PowerShell -File",
                    "host": version["stdout"].strip(), "shell_executable": str(host), "controlled_python": sys.executable,
                    "builtin_module_path": environment["PSModulePath"], "python_path_first": str(Path(sys.executable).parent),
                    "oracle_id": oracle_id, "outputs": actual, "input_unchanged": True, "cwd_engine_untouched": True,
                    "owned_neighbor_unchanged": True, "run_id": execution["run_id"], "manifest_validated": True,
                    "output_location": "<observed-Documents>/<owned-input-GUID>_<timestamp>_<run-GUID>",
                    **{key: value.replace(str(documents), "<observed-Documents>") if isinstance(value, str) else value
                       for key, value in result.items()}}
        finally:
            if cleanup_safe:
                if final is not None and members is not None:
                    remove_known_directory(final, documents, members, ".WinBookSplit-owner.json",
                                           {"schema_version": 1, "kind": "run", "run_id": execution["run_id"]})
                if log is not None and log_validated:
                    remove_known_directory(log.parent, documents, {".WinBookSplit-console-owner.json", "console.log"},
                                           ".WinBookSplit-console-owner.json", owner)
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
