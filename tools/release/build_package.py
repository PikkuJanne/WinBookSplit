#!/usr/bin/env python3
"""Build or verify an unpublished, allowlisted package from an explicit Git commit."""
from __future__ import annotations

import argparse
from hashlib import sha256
import importlib.util
import io
import json
import os
from pathlib import Path
import stat
import struct
import sys
import zipfile


ROOT = Path(__file__).resolve().parents[2]
BUILDER_PATHS = ("tools/release/build_package.py", "tools/release/payload.json",
                 "tools/ci/check_package_inputs.py", "tools/release/frozen_assets.py")
SPEC = importlib.util.spec_from_file_location("wbs_package_inputs", ROOT / BUILDER_PATHS[2])
checker = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(checker)
ASSET_SPEC = importlib.util.spec_from_file_location("wbs_frozen_assets", ROOT / BUILDER_PATHS[3])
assets = importlib.util.module_from_spec(ASSET_SPEC)
ASSET_SPEC.loader.exec_module(assets)
VERSION = "1.0.2"
ZIP_NAME = "WinBookSplit-v1.0.2.zip"
ASSET_NAMES = (ZIP_NAME, "release-manifest.json", "SHA256SUMS.txt")
INCOMPLETE = ".package-incomplete"
TIMESTAMP = assets.TIMESTAMP
FILE_ATTRIBUTES = assets.FILE_ATTRIBUTES
MAX_ARCHIVE_BYTES = assets.MAX_ARCHIVE_BYTES
MAX_MANIFEST_BYTES = assets.MAX_MANIFEST_BYTES
PackageError = assets.PackageError
canonical = assets.canonical
_ordinary = assets.ordinary


def _destination(repo: Path, output: Path, *, existing: bool) -> Path:
    output = Path(output)
    if not output.is_absolute() or output.drive.startswith("\\\\") \
            or any(part in {".", ".."} for part in output.parts) \
            or not checker.safe_relative(output.name):
        raise PackageError("absolute_output_required")
    if output.is_relative_to(repo.resolve(strict=True)):
        raise PackageError("external_output_required")
    # Do not follow junctions/symlinks in any supplied ancestor, even outside Git.
    for parent in reversed(output.parents):
        _ordinary(parent, directory=True)
    if existing:
        _ordinary(output, directory=True)
    elif output.exists() or output.is_symlink():
        raise PackageError("output_already_exists")
    return output


def _blob(repo: Path, commit: str, path: str) -> tuple[bytes, dict]:
    records = checker.git(repo, "ls-tree", "-r", "-z", commit, "--", path).split(b"\0")
    if len(records) != 2 or not records[0]:
        raise PackageError("missing_builder_input")
    metadata, name = records[0].split(b"\t", 1)
    mode, kind, oid = metadata.decode("ascii").split(" ")
    if name.decode("utf-8") != path or mode not in {"100644", "100755"} \
            or kind != "blob" or not checker.SHA1.fullmatch(oid):
        raise PackageError("invalid_builder_input")
    size = int(checker.git(repo, "cat-file", "-s", oid))
    if not 0 < size <= checker.MAX_FILE_BYTES:
        raise PackageError("invalid_builder_size")
    data = checker.git(repo, "cat-file", "blob", oid)
    if len(data) != size:
        raise PackageError("builder_blob_size_mismatch")
    return data, {"path": path, "git_blob": oid, "git_mode": mode,
                  "bytes": size, "sha256": sha256(data).hexdigest()}


def _policy_metadata(payload: dict[str, bytes]) -> dict:
    # Only execute the already Git-blob-validated shipped validator, never a
    # local sibling or a helper from an untrusted asset manifest. The frozen
    # consumer remains pure stdlib and never executes application payload code.
    path = "engine/winbooksplit_compatibility.py"
    namespace = {"__name__": "wbs_bound_release_policy", "__file__": path}
    try:
        exec(compile(payload[path], path, "exec"), namespace)
        policy = json.loads(payload["engine/WinBookSplit.Compatibility.json"].decode("utf-8"),
                            object_pairs_hook=namespace["_unique_object"],
                            parse_constant=namespace["_invalid_constant"])
        namespace["_validate_policy"](policy)
        return assets.dependency_metadata(payload["engine/WinBookSplit.Compatibility.json"],
                                          payload["requirements.txt"])
    except (KeyError, ValueError, TypeError, UnicodeError, RecursionError) as error:
        raise PackageError("invalid_dependency_policy") from error


def _source(repo: Path, commit: str) -> tuple[dict, dict[str, bytes], dict]:
    inputs = checker.validate_inputs(repo, commit)
    if inputs["application_version"] != VERSION:
        raise PackageError("release_version_mismatch")
    builder_files = []
    allowlist = None
    for path in BUILDER_PATHS:
        data, row = _blob(repo, commit, path)
        local = ROOT / path
        _ordinary(local)
        if local.read_bytes() != data:
            raise PackageError("builder_source_mismatch")
        builder_files.append(row)
        if path.endswith("payload.json"):
            allowlist = json.loads(data)
    expected = sorted(set(checker.INPUT_PATHS) - {"docs/codex-v1.0.0/SUPPORT_AND_SETUP.md"})
    if not isinstance(allowlist, dict) or type(allowlist.get("schema_version")) is not int \
            or allowlist != {"schema_version": 1, "version": VERSION, "paths": expected}:
        raise PackageError("invalid_payload_allowlist")
    files = [row for row in inputs["files"] if row["path"] in expected]
    payload = {}
    for row in files:
        data = checker.git(repo, "cat-file", "blob", row["git_blob"])
        if len(data) != row["bytes"] or sha256(data).hexdigest() != row["sha256"]:
            raise PackageError("source_blob_mismatch")
        payload[row["path"]] = data
    # Every shipped relative Markdown link must resolve inside the payload.
    # Developer/historical references use explicit source-repository URLs.
    import re
    for path, data in payload.items():
        if not path.endswith(".md"):
            continue
        for match in re.finditer(r"\]\(([^)]+)\)", data.decode("utf-8-sig")):
            target = match.group(1).split("#", 1)[0]
            if not target or "://" in target:
                continue
            resolved = os.path.normpath(str(Path(path).parent / target)).replace("\\", "/")
            if resolved not in payload:
                raise PackageError("unshipped_document_link")
    tool = {"version": "1", "files": builder_files,
            "python": sys.version.split()[0],
            "git": checker.git(repo, "--version").decode("ascii").strip(),
            "archive_recipe": {"compression": "ZIP_STORED", "timestamp": list(TIMESTAMP),
                               "order": "path Unicode ascending", "mode": "100644",
                               "creator_system": "Unix", "directory_entries": False}}
    base = {"schema_version": 1, "application": "WinBookSplit", "version": VERSION,
            "source_commit": commit, "source_tree": inputs["source_tree"],
            "payload_sha256": sha256(canonical(files)).hexdigest(), "files": files,
            "runtime_sibling_dependencies": inputs["runtime_sibling_dependencies"],
            "dependencies": _policy_metadata(payload),
            "builder": tool, "signed": False,
            "scope": "Unpublished package provenance; Windows/Explorer/Calibre acceptance and release gates remain separate."}
    return base, payload, tool


def _manifest(base: dict, archive: bytes, tool: dict | None = None) -> dict:
    result = dict(base)
    if tool is not None:
        result["builder"] = tool
    result["archive"] = {"name": ZIP_NAME, "bytes": len(archive), "sha256": sha256(archive).hexdigest()}
    return result


def _sums(archive: bytes, manifest: bytes) -> bytes:
    return (sha256(archive).hexdigest() + "  " + ZIP_NAME + "\n" +
            sha256(manifest).hexdigest() + "  release-manifest.json\n").encode("ascii")


_read = assets.read_bounded


def _validate(base: dict, payload: dict[str, bytes], output: Path, *, building: bool = False) -> dict:
    expected_names = set(ASSET_NAMES) | ({INCOMPLETE} if building else set())
    if {path.name for path in output.iterdir()} != expected_names:
        raise PackageError("asset_membership_mismatch")
    archive = _read(output / ZIP_NAME, MAX_ARCHIVE_BYTES)
    manifest_bytes = _read(output / "release-manifest.json", MAX_MANIFEST_BYTES)
    sums = _read(output / "SHA256SUMS.txt", 512)
    try:
        manifest = json.loads(manifest_bytes)
        # Tool runtime may differ when verifying on another supported machine;
        # keep the recorded build runtime, while requiring the exact bound tools/recipe.
        recorded_tool = manifest["builder"]
        expected_tool = dict(base["builder"])
        for name in ("python", "git"):
            value = recorded_tool[name]
            if not isinstance(value, str) or not 0 < len(value) <= 100 \
                    or any(ord(char) < 32 or ord(char) > 126 for char in value):
                raise PackageError("invalid_builder_runtime")
            expected_tool[name] = value
        expected_manifest = _manifest(base, archive, expected_tool)
        if manifest_bytes != canonical(expected_manifest):
            raise PackageError("manifest_mismatch")
        if sums != _sums(archive, manifest_bytes):
            raise PackageError("checksum_mismatch")
        assets.validate_archive(archive, payload)
    except (KeyError, TypeError, json.JSONDecodeError, zipfile.BadZipFile, RuntimeError, NotImplementedError) as error:
        raise PackageError("invalid_package") from error
    return manifest


def validate_package(repo: Path, commit: str, output: Path) -> dict:
    base, payload, _ = _source(Path(repo), commit)
    destination = _destination(Path(repo), output, existing=True)
    return _validate(base, payload, destination)


_archive = assets.archive_bytes


def build_package(repo: Path, commit: str, output: Path) -> dict:
    base, payload, _ = _source(Path(repo), commit)
    destination = _destination(Path(repo), output, existing=False)
    if sum(map(len, payload.values())) > MAX_ARCHIVE_BYTES - 1024 * 1024:
        raise PackageError("payload_too_large")
    destination.mkdir()  # Exclusive; never reuse or recursively clean any destination.
    with (destination / INCOMPLETE).open("xb") as stream:
        stream.write(b"Incomplete build; do not distribute. Retained for diagnosis.\n")
    with (destination / ZIP_NAME).open("xb") as stream:
        stream.write(_archive(payload))
    archive_bytes = _read(destination / ZIP_NAME, MAX_ARCHIVE_BYTES)
    manifest = _manifest(base, archive_bytes)
    manifest_bytes = canonical(manifest)
    with (destination / "release-manifest.json").open("xb") as stream:
        stream.write(manifest_bytes)
    with (destination / "SHA256SUMS.txt").open("xb") as stream:
        stream.write(_sums(archive_bytes, manifest_bytes))
    result = _validate(base, payload, destination, building=True)
    (destination / INCOMPLETE).unlink()  # The only deletion: this run's own fixed marker.
    return result


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("build", "verify"))
    parser.add_argument("--repo", type=Path, required=True)
    parser.add_argument("--commit", required=True, help="Full lowercase 40-digit commit SHA, never HEAD/ref")
    parser.add_argument("--output", type=Path, required=True, help="Absolute external asset directory")
    arguments = parser.parse_args(argv)
    try:
        operation = build_package if arguments.action == "build" else validate_package
        result = operation(arguments.repo, arguments.commit, arguments.output)
        print("PACKAGE_PASSED: " + result["archive"]["sha256"])
        return 0
    except (PackageError, checker.InputError) as error:
        print("PACKAGE_FAILED: " + error.code, file=sys.stderr)
    except (OSError, UnicodeError, ValueError, zipfile.BadZipFile):
        print("PACKAGE_FAILED: invalid_source_or_destination", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
