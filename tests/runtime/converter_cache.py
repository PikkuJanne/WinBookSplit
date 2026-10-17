"""Test-only converter-absence cache observation; never runtime selection input."""
from __future__ import annotations

from contextlib import ExitStack
from hashlib import sha256
import json
import math
import os
from pathlib import Path
import re
import stat

WINDOWS_DIRS = ("Microsoft", "Microsoft/Windows", "Microsoft/Windows/Caches")
MANAGER_DIRS = ("Python", "Python/_cache")
MARKER = "Python/_cache/last_welcome.txt"
MAX_MARKER_BYTES = 4096


def need(ok, message):
    if not ok:
        raise ValueError("Converter absence cache: " + message)


def welcome_bytes(version):
    need(type(version) is str and re.fullmatch(r"[0-9]{1,6}\.[0-9]{1,6}", version) is not None,
         "invalid informational welcome version")
    return (version + "\r\n\r\n!B!Python install manager was successfully updated to " + version +
            ".!W!\r\n\r\nPlease see !B!https://github.com/python/pymanager/releases/tag/" + version +
            "!W! for all changes.\r\n").encode("ascii")


def ordinary(path, directory):
    info = path.lstat()
    need(not stat.S_ISLNK(info.st_mode) and not getattr(info, "st_file_attributes", 0) & 0x400 and
         (stat.S_ISDIR(info.st_mode) if directory else stat.S_ISREG(info.st_mode)), "nonordinary path")
    if not directory:
        need(info.st_nlink == 1, "linked informational marker")
    return info


def identity(info):
    return {"device": info.st_dev, "inode": info.st_ino}


def names(path, maximum):
    result = set()
    with os.scandir(path) as entries:
        for entry in entries:
            result.add(entry.name)
            need(len(result) <= maximum, "unexpected file or directory closure")
    return result


def children(path, expected):
    need(names(path, len(expected)) == set(expected), "unexpected file or directory closure")


def observe(root, windows=None):
    """Keep Windows caches empty; describe only one optional manager marker."""
    root = Path(root)
    ordinary(root, True)
    root_names = names(root, 2)
    need(root_names in ({"Microsoft"}, {"Microsoft", "Python"}), "unexpected isolation-root member")
    result = {}
    for relative, expected in zip(WINDOWS_DIRS, (("Windows",), ("Caches",), ())):
        path = root / relative
        result[relative] = identity(ordinary(path, True))
        children(path, expected)
    marker = None
    if "Python" in root_names:
        with ExitStack() as leases:
            directories = {}
            for relative, expected in zip(MANAGER_DIRS, (("_cache",), ("last_welcome.txt",))):
                path = root / relative
                directories[relative] = identity(ordinary(path, True))
                if windows is not None:
                    leases.enter_context(windows.DirectoryGuard(path))
                children(path, expected)
            path = root / MARKER
            before = ordinary(path, False)
            need(0 < before.st_size <= MAX_MARKER_BYTES, "informational marker size is unbounded")
            if windows is not None:
                lease = leases.enter_context(windows.FileGuard(path))
            with path.open("rb") as stream:
                opened = os.fstat(stream.fileno())
                raw = stream.read(MAX_MARKER_BYTES + 1)
            after = ordinary(path, False)
            need(identity(before) == identity(opened) == identity(after) and
                 before.st_size == after.st_size == len(raw), "informational marker changed during observation")
            try:
                version = raw.split(b"\r\n", 1)[0].decode("ascii", errors="strict")
            except UnicodeError as error:
                raise ValueError("Converter absence cache: malformed welcome marker") from error
            need(raw == welcome_bytes(version), "unknown informational marker content")
            if windows is not None:
                lease.assert_unchanged()
            marker = {"directories": directories, "relative": MARKER, "bytes": len(raw),
                      "sha256": sha256(raw).hexdigest(), "version": version, "links": 1,
                      **identity(after)}
    return {"windows_directories": result, "manager_marker": marker}


def _identity(value):
    need(type(value) is dict and set(value) == {"device", "inode"} and
         type(value["device"]) is int and value["device"] >= 0 and
         type(value["inode"]) is int and value["inode"] > 0, "invalid directory identity")


def validate_marker(marker):
    need(type(marker) is dict and set(marker) == {"directories", "relative", "bytes", "sha256", "version", "links", "device", "inode"},
         "invalid manager marker schema")
    need(type(marker["directories"]) is dict and set(marker["directories"]) == set(MANAGER_DIRS), "manager directory closure differs")
    for item in marker["directories"].values():
        _identity(item)
    _identity({key: marker[key] for key in ("device", "inode")})
    raw = welcome_bytes(marker["version"])
    need(marker["relative"] == MARKER and type(marker["bytes"]) is int and
         marker["bytes"] == len(raw) <= MAX_MARKER_BYTES and marker["links"] == 1 and
         type(marker["links"]) is int and marker["sha256"] == sha256(raw).hexdigest(), "unproved welcome marker identity")


def validate_reference(reference, observed):
    need(type(reference) is dict and reference.get("kind") == "VERIFIED_INSTALLED_ONLY_MANAGER_CACHE_REFERENCE" and
         reference.get("parent_environment_unchanged") is True and reference.get("runtime_acquisition_attempted") is False and
         reference.get("supplied_to_application") is False, "manager reference purpose or preservation differs")
    rows = reference.get("probes")
    need(type(rows) is list and 1 <= len(rows) <= 2, "bounded official manager reference missing")
    matched = False
    for row in rows:
        need(type(row) is dict and row.get("listing_command") == ["list", "--format=exe"] and
             row.get("child_localappdata_only") is True and row.get("parent_localappdata_restored") is True and
             row.get("marker_absent_before") is True, "reference is not installed-only cache observation")
        need(type(row.get("registration")) is dict and type(row["registration"].get("Package")) is dict,
             "manager registration is missing")
        registration = row["registration"]["Package"]
        tuples = {
            ("PythonSoftwareFoundation.PythonManager_qbz5n2kfra8p0", "CN=4975D53F-AA7E-49A5-8B49-EA4FDC1BB66B", "Store"),
            ("PythonSoftwareFoundation.PythonManager_3847v3x7pw1km", "CN=Python Software Foundation, O=Python Software Foundation, L=Beaverton, S=Oregon, C=US", "Developer"),
        }
        need(registration.get("Name") == "PythonSoftwareFoundation.PythonManager" and
             all(type(registration.get(key)) is str for key in ("PackageFamilyName", "Publisher", "SignatureKind")) and
             (registration.get("PackageFamilyName"), registration.get("Publisher"), registration.get("SignatureKind")) in tuples and
             registration.get("Architecture") == "X64" and registration.get("Status") == "Ok" and
             all(registration.get(key) is False for key in ("IsDevelopmentMode", "IsPartiallyStaged", "IsFramework", "IsResourcePackage")),
             "unverified manager registration tuple")
        version = registration.get("Version")
        need(type(version) is str and re.fullmatch(r"[0-9]{1,6}(?:\.[0-9]{1,6}){3}", version) is not None,
             "manager package version is missing")
        package_full_name = (registration["Name"] + "_" + version + "_x64__" +
                             registration["PackageFamilyName"].rsplit("_", 1)[1])
        need(registration.get("PackageFullName") == package_full_name, "manager full package identity is inconsistent")
        digest = row.get("manager_sha256_before")
        need(type(digest) is str and re.fullmatch(r"[0-9a-f]{64}", digest) is not None and digest == row.get("manager_sha256_after"),
             "manager image changed or unproved")
        probe = row.get("listing")
        need(type(probe) is dict, "native manager listing receipt is missing")
        need(type(probe.get("Pid")) is int and probe["Pid"] > 0 and type(probe.get("ExitCode")) is int and probe["ExitCode"] == 0 and
             probe.get("PackageIdentityVerified") is True and probe.get("PackageFullName") == package_full_name and
             all(probe.get(key) is True for key in ("JobAssigned", "ParentStopped", "DescendantsStopped", "StreamsComplete", "InputWriterStopped")) and
             all(probe.get(key) is False for key in ("TimedOut", "Cancelled", "StdoutTruncated", "StderrTruncated")) and
             all(key in probe and probe[key] is None for key in ("StartError", "StopError", "StreamError", "ResultError", "InteractionError", "InputError")),
             "bound manager listing did not complete safely")
        elapsed = probe.get("ElapsedSeconds")
        need(type(elapsed) in (int, float) and math.isfinite(elapsed) and 0 <= elapsed <= 13, "manager reference process duration differs")
        for stream in ("Stdout", "Stderr"):
            need(type(probe.get(stream)) is str and type(probe.get(stream + "TotalBytes")) is int and
                 probe[stream + "TotalBytes"] == len(probe[stream].encode("utf-8")) <= 65536, "manager reference streams differ")
        marker = row.get("manager_marker")
        if marker is not None:
            validate_marker(marker)
            matched |= all(marker[key] == observed[key] for key in ("relative", "bytes", "sha256", "version"))
    need(matched, "observed welcome marker does not match verified manager reference bytes")


def validate_environment(record):
    need(type(record) is dict and set(record) == {"schema_version", "child_local_app_data_isolated", "windows_cache_inventory_unchanged",
         "ordinary_directory_guards_unchanged", "cache_directory_identities_before", "cache_directory_identities_after",
         "parent_local_app_data_unchanged", "fixed_system_calibre_paths_absent", "manager_cache_observation"}, "environment schema differs")
    need(type(record["schema_version"]) is int and record["schema_version"] == 2, "unsupported cache observation schema")
    need(all(record[key] is True for key in ("child_local_app_data_isolated", "windows_cache_inventory_unchanged",
         "ordinary_directory_guards_unchanged", "parent_local_app_data_unchanged", "fixed_system_calibre_paths_absent")), "isolation or Windows cache preservation is unproved")
    before, after = record["cache_directory_identities_before"], record["cache_directory_identities_after"]
    need(type(before) is dict and type(after) is dict and set(before) == set(after) == set(WINDOWS_DIRS) and before == after,
         "Windows cache closure or identities changed")
    for item in before.values():
        _identity(item)
    manager = record["manager_cache_observation"]
    need(type(manager) is dict and set(manager) == {"before", "after", "reference"} and manager["before"] is None,
         "manager cache was prebuilt or observation is missing")
    if manager["after"] is None:
        need(manager["reference"] is None, "unneeded manager reference is contradictory")
    else:
        validate_marker(manager["after"])
        validate_reference(manager["reference"], manager["after"])


def reference(root, repository, inbox, history, windows):
    """A separate bounded developer probe, never setup/application parameters."""
    work = Path(root).parent / "manager-cache-reference"
    work.mkdir()
    report = work / "reference.json"
    script = Path(repository) / "tests/runtime/Observe-ManagerCache.ps1"
    command = [str(inbox), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-File", str(script),
               "-RepositoryRoot", str(repository), "-WorkRoot", str(work), "-ReportPath", str(report)]
    result = history.run_entrypoint(command, work, environment=history.clean_environment(work), stdin="", timeout=60)
    need(result["exit_code"] == 0, "manager reference controller failed")
    ordinary(report, False)
    need(0 < report.stat().st_size <= 1024 * 1024, "manager reference receipt is unbounded")
    value = json.loads(report.read_text(encoding="utf-8"))
    rows = value.get("probes")
    need(type(rows) is list and 1 <= len(rows) <= 2, "bounded official manager reference missing")
    for row in rows:
        cache = Path(row.pop("cache_root"))
        need(cache.parent == work, "manager reference cache escaped owned work")
        row["manager_marker"] = observe(cache, windows)["manager_marker"]
    value["controller"] = {"command": command, "cwd": str(work), **result}
    return value


def environment_record(before, after, provenance):
    record = {"schema_version": 2, "child_local_app_data_isolated": True, "windows_cache_inventory_unchanged": True,
              "ordinary_directory_guards_unchanged": True,
              "cache_directory_identities_before": before["windows_directories"],
              "cache_directory_identities_after": after["windows_directories"],
              "parent_local_app_data_unchanged": True, "fixed_system_calibre_paths_absent": True,
              "manager_cache_observation": {"before": before["manager_marker"], "after": after["manager_marker"], "reference": provenance}}
    validate_environment(record)
    return record
