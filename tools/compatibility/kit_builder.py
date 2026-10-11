"""Build the finite owner kit, or verify/extract its externally pinned release.

The owner branch needs only regular CPython's standard library. It imports the
two already authenticated kit helpers and extracts the verifier's returned
bytes once under the shipped Windows guards. It never invokes Git or pip.
"""
from __future__ import annotations

import argparse
from contextlib import ExitStack, nullcontext
from hashlib import sha256
import importlib.util
from io import BytesIO
import json
import os
from pathlib import Path
import re
import stat
import subprocess
import sys
import zipfile

MAX_BINDINGS = 128 * 1024
MAX_CODE = 2 * 1024 * 1024
MAX_FIXTURE = 8 * 1024 * 1024
MAX_KIT = 64 * 1024 * 1024
HEX = re.compile(r"[0-9a-f]{64}\Z")
BOOT = ("engine/WinBookSplit.Logging.ps1", "engine/WinBookSplit.Diagnostics.ps1",
        "engine/WinBookSplit.Process.ps1", "engine/WinBookSplit.Support.ps1",
        "tools/compatibility/Invoke-NativeAcceptance.ps1")
FIXTURES = ("flat-6.pdf", "nested-12.pdf", "original-three-chapters.azw3",
            "original-three-chapters.epub", "simple-10.pdf", "native-fixtures.json")
PREREQUISITES = {
    "python": {"name": "python-3.14.8-amd64.zip", "version": "3.14.8", "bytes": 32879974,
               "url": "https://www.python.org/ftp/python/3.14.8/python-3.14.8-amd64.zip",
               "sha256": "4873947a8afc037846b180312b83c744a4146a851cfd316a75c3125a4d8299da"},
    "pypdf": {"name": "pypdf-6.20.0-py3-none-any.whl", "version": "6.20.0", "bytes": 401710,
              "url": "https://files.pythonhosted.org/packages/18/42/a945f65cc61c739ec80f4112c4b78ed1791f25d33f45f19389c9c9e247e2/pypdf-6.20.0-py3-none-any.whl",
              "sha256": "f003fc2014814d264fe7dd3f9d435c158e23e1a85a2233f87a0a2d6d21c914ad"},
    "calibre": {"name": "calibre-portable-installer-9.15.0.exe", "version": "9.15.0", "bytes": 205319944,
                "url": "https://download.calibre-ebook.com/9.15.0/calibre-portable-installer-9.15.0.exe",
                "sha256": "55a8c89bc0739a2dc6d496742ea625fccc6dfdec1413eb805805a28e7227536c"}}


class KitError(ValueError):
    pass


def need(condition, code):
    if not condition:
        raise KitError(code)


def canonical(value):
    return (json.dumps(value, sort_keys=True, indent=2, ensure_ascii=True) + "\n").encode("ascii")


def unique(raw):
    def pairs(rows):
        output, seen = {}, set()
        for key, value in rows:
            need(key.casefold() not in seen, "duplicate_json_key")
            seen.add(key.casefold())
            output[key] = value
        return output
    return json.loads(raw.decode("utf-8", "strict"), object_pairs_hook=pairs,
                      parse_constant=lambda _: (_ for _ in ()).throw(KitError("nonfinite_json")))


def ordinary(path, *, directory=False):
    path = Path(path)
    need(path.is_absolute() and Path(os.path.abspath(path)) == path, "literal_absolute_path_required")
    if os.name == "nt":
        need(re.fullmatch(r"[A-Za-z]:", path.drive) is not None, "local_drive_path_required")
    for entry in [*reversed(path.parents), path]:
        info = entry.lstat()
        need(not stat.S_ISLNK(info.st_mode) and not getattr(info, "st_file_attributes", 0) & 0x400,
             "reparse_path_rejected")
        need(stat.S_ISDIR(info.st_mode) if entry != path or directory else stat.S_ISREG(info.st_mode),
             "ordinary_path_required")
    if not directory:
        need(path.stat().st_nlink == 1, "hardlink_file_rejected")
    return path


def bounded(path, maximum):
    path = ordinary(path)
    need(0 < path.stat().st_size <= maximum, "file_size_invalid")
    with path.open("rb") as stream:
        data = stream.read(maximum + 1)
    need(0 < len(data) <= maximum, "file_size_invalid")
    return data


def load_authenticated(path, row, name):
    data = bounded(path, MAX_CODE)
    need(type(row) is dict and set(row) == {"bytes", "sha256"}
         and type(row["bytes"]) is int and row["bytes"] == len(data)
         and HEX.fullmatch(row["sha256"]) and sha256(data).hexdigest() == row["sha256"],
         "helper_identity_mismatch")
    # The central native harness retains a strict file lease during this call.
    module = importlib.util.module_from_spec(importlib.util.spec_from_loader(name, loader=None))
    module.__file__ = str(path)
    sys.modules[name] = module
    exec(compile(data, str(path), "exec"), module.__dict__)
    return module


def rows(payload):
    return {name: {"bytes": len(data), "sha256": sha256(data).hexdigest()}
            for name, data in sorted(payload.items())}


def aggregate(mapping):
    return sha256(canonical(mapping)).hexdigest()


def guarded_write_tree(target, payload, windows, *, held_stack=None):
    need(type(payload) is dict and 1 <= len(payload) <= 128, "payload_count_invalid")
    names, directories = set(), {}
    for name, data in payload.items():
        need(type(name) is str and type(data) is bytes and 0 < len(data) <= 32 * 1024 * 1024,
             "payload_member_invalid")
        parts = name.split("/")
        need(not name.startswith("/") and all(part not in {"", ".", ".."} and part == part.rstrip(" .")
             and not re.fullmatch(r"(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?", part, re.I)
             and not any(ord(c) < 32 or ord(c) == 127 or c in '<>:"\\|?*' for c in part) for part in parts),
             "payload_path_invalid")
        need(name.casefold() not in names, "payload_path_collision")
        names.add(name.casefold())
        for length in range(1, len(parts)):
            directory = "/".join(parts[:length])
            need(directory.casefold() not in directories or directories[directory.casefold()] == directory,
                 "payload_ancestor_spelling_collision")
            directories[directory.casefold()] = directory
    need(not names.intersection(directories) and sum(map(len, payload.values())) <= MAX_KIT,
         "payload_ancestor_or_size_collision")
    ordinary(target.parent, directory=True)
    with (ExitStack() if held_stack is None else nullcontext(held_stack)) as stack:
        for ancestor in [*reversed(target.parent.parents), target.parent]:
            stack.enter_context(windows.DirectoryGuard(ancestor))
        target.mkdir()  # exclusive: an existing file/directory is never adopted.
        guards = {target: stack.enter_context(windows.DirectoryGuard(target))}
        for relative, data in sorted(payload.items()):
            parts = relative.split("/")
            parent = target
            for component in parts[:-1]:
                child = parent / component
                if child not in guards:
                    guards[parent].assert_unchanged()
                    child.mkdir()
                    guards[child] = stack.enter_context(windows.DirectoryGuard(child))
                parent = child
            guards[parent].assert_unchanged()
            destination = parent / parts[-1]
            with destination.open("xb") as stream:
                stream.write(data)
                stream.flush()
                os.fsync(stream.fileno())
            ordinary(destination)
            with windows.FileGuard(destination) as guard:
                observed = bounded(destination, 32 * 1024 * 1024)
                need(observed == data, "extracted_bytes_differ")
                guard.assert_unchanged()
        for guard in guards.values():
            guard.assert_unchanged()
        return guards[target] if held_stack is not None else None


def verify_extract(args):
    need(sys.flags.isolated and sys.dont_write_bytecode and os.name == "nt", "isolated_windows_python_required")
    root = ordinary(args.kit_root, directory=True)
    raw = bounded(root / "kit-bindings.json", MAX_BINDINGS)
    need(HEX.fullmatch(args.bindings_sha256) and sha256(raw).hexdigest() == args.bindings_sha256,
         "bindings_identity_mismatch")
    bindings = unique(raw)
    need(set(bindings) == {"protocol", "schema_version", "production", "kit", "helpers", "fixtures", "prerequisites"}
         and bindings["protocol"] == "winbooksplit.native-kit" and type(bindings["schema_version"]) is int
         and bindings["schema_version"] == 1, "bindings_schema_invalid")
    helpers = bindings["helpers"]
    need(bindings["kit"]["helpers_payload_sha256"] == aggregate(helpers)
         and bindings["kit"]["fixtures_payload_sha256"] == aggregate(bindings["fixtures"]),
         "kit_aggregate_identity_mismatch")
    verifier = load_authenticated(root / "tools/release/frozen_assets.py", helpers["tools/release/frozen_assets.py"], "wbs_frozen_assets")
    windows = load_authenticated(root / "engine/winbooksplit_windows.py", helpers["engine/winbooksplit_windows.py"], "wbs_kit_windows")
    with ExitStack() as stack:
        for ancestor in [*reversed(root.parents), root]:
            stack.enter_context(windows.DirectoryGuard(ancestor))
        for relative in ("tools/release/frozen_assets.py", "engine/winbooksplit_windows.py", "kit-bindings.json"):
            stack.enter_context(windows.FileGuard(root / relative))
        manifest, payload = verifier.verify_frozen_assets(root / "assets", bindings["production"])
        guarded_write_tree(args.destination, payload, windows)
        record = {"protocol": "winbooksplit.native-extraction", "schema_version": 1, "success": True,
                  "production": bindings["production"], "files": rows(payload),
                  "destination": str(args.destination), "payload_sha256": manifest["payload_sha256"]}
        need(args.output.parent == args.destination.parent, "receipt_parent_must_be_owned_work")
        for ancestor in [*reversed(args.output.parent.parents), args.output.parent]:
            stack.enter_context(windows.DirectoryGuard(ordinary(ancestor, directory=True)))
        with args.output.open("xb") as stream:
            stream.write(canonical(record))
    return record


def powershell_quote(value):
    return "'" + value.replace("'", "''") + "'"


def wrapper(bindings_digest, helper_rows, logging_source, production):
    # Embed the already existing native lease definition, byte-for-byte. This
    # keeps pre-import authentication independent of an unverified helper file.
    logging_text = logging_source.decode("utf-8-sig").replace("\r\n", "\n")
    native_definition = logging_text.split("Add-Type -TypeDefinition @'\n", 1)[1].split("\n'@", 1)[0]
    bootstrap_report = {"protocol": "winbooksplit.native-acceptance", "schema_version": 1,
        "evidence_kind": "NATIVE_AUTOMATED", "execution": "__EXECUTION__", "started_utc": "__UTC__", "finished_utc": "__UTC__",
        "success": False, "os": None, "oracle": None, "identities": None, "expected_production": production,
        "cases": [{"id": name, "status": "NOT_RUN", "exit_code": None, "error_code": "bootstrap_failed"}
          for name in ("ref-epub", "ref-azw3", "manual", "manual-first", "auto1", "auto2", "repeat", "preview", "invalid-manual", "no-bookmarks", "cancel-plan", "missing-python", "missing-pypdf", "epub-discard", "epub-keep", "azw3-discard", "azw3-keep", "bat-one", "bat-two")],
        "stages": [{"id": "bootstrap", "status": "FAIL", "code": "bootstrap_failed", "exit_code": None, "probe_sha256": None}],
        "controls": [], "cleanup_safe": False}
    import base64
    early_json = base64.b64encode(canonical(bootstrap_report)).decode("ascii")
    code = ["# One-session finite native kit. Generated source is pinned by the delivered kit SHA.",
            "[CmdletBinding()]", "param([string]$BarePythonPath,[string]$CalibrePath,[string]$Wheelhouse,",
            "      [switch]$ProvisionPrerequisites,[switch]$AllowWindows11DryRun,",
            "      [ValidateSet('AGENT_LAUNCHED','USER_LAUNCHED')][string]$Execution='USER_LAUNCHED')",
            "$ErrorActionPreference='Stop'; $PSModuleAutoloadingPreference='None'",
            "Import-Module ([IO.Path]::Combine($PSHOME,'Modules\\Microsoft.PowerShell.Utility\\Microsoft.PowerShell.Utility.psd1')) -ErrorAction Stop",
            "Import-Module ([IO.Path]::Combine($PSHOME,'Modules\\Microsoft.PowerShell.Management\\Microsoft.PowerShell.Management.psd1')) -ErrorAction Stop",
            "Add-Type -TypeDefinition @'", native_definition, "'@ -ErrorAction Stop",
            "$leases=New-Object 'Collections.Generic.List[IDisposable]'",
            "try {", "  $root=[IO.Path]::GetFullPath($PSScriptRoot)",
            "  $ancestors=New-Object 'Collections.Generic.List[string]'; $current=$root",
            "  while ($current) { $ancestors.Add($current); $current=[IO.Path]::GetDirectoryName($current) }",
            "  for ($i=$ancestors.Count-1;$i -ge 0;$i--) { $leases.Add((New-Object WinBookSplitLogging.PathLease($ancestors[$i],$true))) }",
            "  $expected=@{" ]
    for name in ("kit-bindings.json", "engine/WinBookSplit.Outcomes.json", *BOOT):
        row = {"sha256": bindings_digest, "bytes": MAX_BINDINGS} if name == "kit-bindings.json" else helper_rows[name]
        code.append(f"    {powershell_quote(name)}=@{{sha256={powershell_quote(row['sha256'])};bytes={row['bytes']}}}")
    code += ["  }", "  foreach ($relative in $expected.Keys) {",
             "    $path=[IO.Path]::Combine($root,$relative.Replace('/','\\'))",
             "    $lease=New-Object WinBookSplitLogging.PathLease($path,$false,$false,$true); $leases.Add($lease)",
             "    $stream=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)",
             "    $algorithm=[Security.Cryptography.SHA256]::Create()",
             "    try {",
             "      $lease.AssertSameHandle($stream.SafeFileHandle)",
             "      if ($stream.Length -lt 1 -or $stream.Length -gt $expected[$relative].bytes) { throw 'Native kit bootstrap size mismatch.' }",
             "      $digest=([BitConverter]::ToString($algorithm.ComputeHash($stream))).Replace('-','').ToLowerInvariant()",
             "      if ($digest -cne $expected[$relative].sha256) { throw 'Native kit bootstrap identity mismatch.' }",
             "    } finally { $stream.Dispose(); $algorithm.Dispose() }",
             "  }", "  foreach ($relative in @('engine/WinBookSplit.Logging.ps1','engine/WinBookSplit.Diagnostics.ps1','engine/WinBookSplit.Process.ps1','engine/WinBookSplit.Support.ps1','tools/compatibility/Invoke-NativeAcceptance.ps1')) {",
             "    . ([IO.Path]::Combine($root,$relative.Replace('/','\\')))", "  }",
             f"  $options=@{{KitRoot=$root;BindingsSha256='{bindings_digest}';Execution=$Execution;BarePythonPath=$BarePythonPath;CalibrePath=$CalibrePath;Wheelhouse=$Wheelhouse;ProvisionPrerequisites=$ProvisionPrerequisites;AllowWindows11DryRun=$AllowWindows11DryRun}}",
             "  $result=Invoke-NativeAcceptance @options", "  exit $(if ($result.success) { 0 } else { 1 })",
             "} catch {",
             "  Write-Host 'Native kit bootstrap failed. No acceptance pass was recorded.'",
             "  $earlyLeases=New-Object 'Collections.Generic.List[IDisposable]'; $earlyFile=$earlyZip=$null",
             "  try {",
             "    $parent=[Environment]::GetFolderPath('UserProfile');$parents=New-Object 'Collections.Generic.List[string]';$entry=$parent",
             "    while($entry){$parents.Add($entry);$entry=[IO.Path]::GetDirectoryName($entry)}",
             "    for($i=$parents.Count-1;$i -ge 0;$i--){$earlyLeases.Add((New-Object WinBookSplitLogging.PathLease($parents[$i],$true)))}",
             f"    $body=[Text.Encoding]::ASCII.GetString([Convert]::FromBase64String('{early_json}')).Replace('__UTC__',[DateTime]::UtcNow.ToString('o')).Replace('__EXECUTION__',$Execution)",
             "    [Reflection.Assembly]::Load('System.IO.Compression, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')|Out-Null",
             "    $returnPath=[IO.Path]::Combine($parent,'WBS-native-bootstrap-'+[Guid]::NewGuid().ToString('N')+'.zip')",
             "    $earlyFile=[IO.File]::Open($returnPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)",
             "    $earlyZip=New-Object IO.Compression.ZipArchive($earlyFile,[IO.Compression.ZipArchiveMode]::Create,$false)",
             "    $member=$earlyZip.CreateEntry('native-result.json',[IO.Compression.CompressionLevel]::NoCompression);$writer=$member.Open()",
             "    try{$bytes=[Text.Encoding]::UTF8.GetBytes($body);$writer.Write($bytes,0,$bytes.Length)}finally{$writer.Dispose()}",
             "    Write-Host ('Return only: '+$returnPath)",
             "  }catch{Write-Host 'The guarded bootstrap return ZIP could not be created. Retain this kit; no pass recorded.'}",
             "  finally{if($earlyZip){$earlyZip.Dispose()};if($earlyFile){$earlyFile.Dispose()};for($i=$earlyLeases.Count-1;$i -ge 0;$i--){$earlyLeases[$i].Dispose()}}",
             "  exit 1", "}",
             "finally { for ($i=$leases.Count-1;$i -ge 0;$i--) { $leases[$i].Dispose() } }", ""]
    return "\r\n".join(code).encode("utf-8")


def git_identity(root, checker):
    # Developer-only reuse of the existing bounded, sanitized repository query.
    commit = checker.git(root, "rev-parse", "HEAD").decode("ascii", "strict").strip()
    tree = checker.git(root, "rev-parse", commit + "^{tree}").decode("ascii", "strict").strip()
    need(re.fullmatch(r"[0-9a-f]{40}", commit) and re.fullmatch(r"[0-9a-f]{40}", tree),
         "repository_identity_invalid")
    return commit, tree


def write_build_outputs(output, payload, windows, bindings):
    # Keep the original output-parent and created target identities through all
    # sibling publication and receipt reads/writes. Never adopt a replacement.
    ordinary(output.parent, directory=True)
    with ExitStack() as stack:
        directories = [stack.enter_context(windows.DirectoryGuard(ancestor))
                       for ancestor in [*reversed(output.parent.parents), output.parent]]
        directories.append(guarded_write_tree(output, payload, windows, held_stack=stack))
        archive = output.with_suffix(".zip")
        for guard in directories:
            guard.assert_unchanged()
        # Bind the closed file to exact authored ZIP bytes; a directory lease
        # alone does not prevent unlink/recreate of a child by this account.
        with BytesIO() as buffer:
            with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_STORED, allowZip64=False) as zipped:
                for name, data in sorted(payload.items()):
                    entry = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                    entry.create_system = 3
                    entry.external_attr = (stat.S_IFREG | 0o644) << 16
                    zipped.writestr(entry, data)
            expected_archive = buffer.getvalue()
        need(0 < len(expected_archive) <= MAX_KIT, "kit_archive_size_invalid")
        with archive.open("xb") as stream:
            stream.write(expected_archive)
            stream.flush()
            os.fsync(stream.fileno())
        archive_guard = stack.enter_context(windows.FileGuard(archive))
        archive_bytes = bounded(archive, MAX_KIT)
        need(archive_bytes == expected_archive, "kit_archive_bytes_differ")
        archive_guard.assert_unchanged()
        record = {"protocol": "winbooksplit.native-kit-build", "schema_version": 1, "success": True,
                  "production": bindings["production"], "kit": bindings["kit"],
                  "bindings_sha256": sha256(canonical(bindings)).hexdigest(),
                  "kit_archive": {"bytes": len(archive_bytes), "sha256": sha256(archive_bytes).hexdigest()},
                  "files": rows(payload)}
        for guard in directories:
            guard.assert_unchanged()
        receipt = output.with_suffix(".build.json")
        with receipt.open("xb") as stream:
            stream.write(canonical(record))
            stream.flush()
            os.fsync(stream.fileno())
        with windows.FileGuard(receipt) as guard:
            need(bounded(receipt, MAX_KIT) == canonical(record), "build_receipt_bytes_differ")
            guard.assert_unchanged()
        for guard in directories:
            guard.assert_unchanged()
    return record


def build(args):
    root = ordinary(args.repository_root, directory=True)
    # Owner verify-extract never imports this developer-only Git consumer.
    checker_path = root / "tools/ci/check_package_inputs.py"
    checker_bytes = bounded(checker_path, MAX_CODE)
    checker = load_authenticated(checker_path, rows({"c": checker_bytes})["c"], "wbs_build_checker")
    verifier_bytes = bounded(args.verifier_path, MAX_CODE)
    verifier = load_authenticated(args.verifier_path, rows({"v": verifier_bytes})["v"], "wbs_build_verifier")
    manifest_raw = bounded(args.assets_root / "release-manifest.json", MAX_BINDINGS)
    manifest = unique(manifest_raw)
    version = manifest["version"]
    names = (f"WinBookSplit-v{version}.zip", "release-manifest.json", "SHA256SUMS.txt")
    asset_bytes = {name: bounded(args.assets_root / name, 32 * 1024 * 1024) for name in names}
    production = {"version": version, "source_commit": manifest["source_commit"],
                  "source_tree": manifest["source_tree"], "payload_sha256": manifest["payload_sha256"],
                  "policy_sha256": next(row["sha256"] for row in manifest["files"] if row["path"] == "engine/WinBookSplit.Compatibility.json"), "assets": rows(asset_bytes)}
    verifier.verify_frozen_assets(args.assets_root, production)
    repository_commit, repository_tree = git_identity(root, checker)
    payload = {"assets/" + name: data for name, data in asset_bytes.items()}
    helpers = {name: bounded(root / name, MAX_CODE) for name in BOOT[:-1]}
    helpers.update({"tools/compatibility/Invoke-NativeAcceptance.ps1": bounded(args.core_path, MAX_CODE),
                    "tools/compatibility/kit_builder.py": bounded(Path(__file__).absolute(), MAX_CODE),
                    "tools/compatibility/native_smoke.py": bounded(args.oracle_path, MAX_CODE),
                    "tools/release/frozen_assets.py": verifier_bytes,
                    "engine/winbooksplit_windows.py": bounded(root / "engine/winbooksplit_windows.py", MAX_CODE),
                    "tests/cli/Probe-Help.ps1": bounded(root / "tests/cli/Probe-Help.ps1", MAX_CODE),
                    "tests/process/session_child.py": bounded(root / "tests/process/session_child.py", MAX_CODE)})
    helpers["engine/WinBookSplit.Outcomes.json"] = bounded(root / "engine/WinBookSplit.Outcomes.json", MAX_CODE)
    if not args.draft_rehearsal:
        for name, data in helpers.items():
            blob = checker.git(root, "cat-file", "blob", repository_commit + ":" + name)
            need(blob == data, "kit_helper_not_committed_bytes")
        kit_source = {"source_commit": repository_commit, "source_tree": repository_tree,
                      "mode": "COMMITTED", "repository_observation": None}
    else:
        kit_source = {"source_commit": None, "source_tree": None, "mode": "EXTERNAL_DRAFT_REHEARSAL",
                      "repository_observation": {"commit": repository_commit, "tree": repository_tree}}
    fixtures = {name: bounded(args.fixture_root / name, MAX_FIXTURE) for name in FIXTURES}
    need(set(entry.name for entry in args.fixture_root.iterdir()) == set(FIXTURES), "fixture_closure_invalid")
    bindings = {"protocol": "winbooksplit.native-kit", "schema_version": 1, "production": production,
                "kit": {**kit_source,
                        "helpers_payload_sha256": aggregate(rows(helpers)), "fixtures_payload_sha256": aggregate(rows(fixtures))},
                "helpers": rows(helpers), "fixtures": rows(fixtures), "prerequisites": PREREQUISITES}
    binding_bytes = canonical(bindings)
    digest = sha256(binding_bytes).hexdigest()
    payload.update(helpers)
    payload.update({"fixtures/" + name: data for name, data in fixtures.items()})
    payload["kit-bindings.json"] = binding_bytes
    payload["Run-Windows10.ps1"] = wrapper(digest, bindings["helpers"], helpers["engine/WinBookSplit.Logging.ps1"], production)
    payload["Run-Windows10.cmd"] = ("@echo off\r\nsetlocal DisableDelayedExpansion\r\n"
        '"%SystemRoot%\\System32\\WindowsPowerShell\\v1.0\\powershell.exe" -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0Run-Windows10.ps1" -ProvisionPrerequisites\r\n'
        "exit /b %errorlevel%\r\n").encode("ascii")
    payload["README.txt"] = ("Run Run-Windows10.cmd once on regular Windows 10 22H2 x64.\r\n"
        "It explicitly provisions pinned private tools, tests the exact bundled release and retains raw local work.\r\n"
        "Return only the printed native-result.zip. It excludes personal paths and raw console text.\r\n"
        "Byte-identical BAT forwarding is automated. The Explorer path is unchanged; no repeated manual sign-off is requested.\r\n"
        "Do not edit pins or reuse a result for different production/kit bytes.\r\n").encode("ascii")
    payload["kit-manifest.json"] = canonical({"protocol": "winbooksplit.native-kit-files", "schema_version": 1,
        "bindings_sha256": digest, "files": rows(payload)})
    windows = load_authenticated(root / "engine/winbooksplit_windows.py", bindings["helpers"]["engine/winbooksplit_windows.py"], "wbs_build_windows")
    return write_build_outputs(args.output, payload, windows, bindings)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    actions = parser.add_subparsers(dest="command", required=True)
    owner = actions.add_parser("verify-extract")
    for name in ("kit-root", "destination", "output"):
        owner.add_argument("--" + name, type=Path, required=True)
    owner.add_argument("--bindings-sha256", required=True)
    developer = actions.add_parser("build")
    for name in ("repository-root", "assets-root", "fixture-root", "core-path", "oracle-path", "verifier-path", "output"):
        developer.add_argument("--" + name, type=Path, required=True)
    developer.add_argument("--draft-rehearsal", action="store_true",
                           help="Disclose external/uncommitted helper bytes; never claim a committed kit source identity.")
    args = parser.parse_args(argv)
    try:
        record = verify_extract(args) if args.command == "verify-extract" else build(args)
        print(json.dumps({"protocol": record["protocol"], "success": True}))
        return 0
    except (KitError, OSError, ValueError, KeyError, TypeError, subprocess.SubprocessError):
        print(json.dumps({"protocol": "winbooksplit.native-kit-error", "success": False, "code": "kit_validation_failed"}))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
