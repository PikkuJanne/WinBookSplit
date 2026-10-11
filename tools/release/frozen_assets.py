"""Shared canonical release-asset checks; no Git, extraction or application imports.

Frozen verification requires external trusted byte identities. Manifest hashes
are never adopted as their own authority. Developer builds separately establish
the manifest's source-to-Git binding in build_package.py.
"""
from __future__ import annotations

from hashlib import sha1, sha256
import io
import json
from pathlib import Path, PurePosixPath
import re
import stat
import struct
import zipfile


TIMESTAMP = (1980, 1, 1, 0, 0, 0)
FILE_ATTRIBUTES = (stat.S_IFREG | 0o644) << 16
MAX_ARCHIVE_BYTES = 32 * 1024 * 1024
MAX_MANIFEST_BYTES = 128 * 1024
MAX_FILE_BYTES = 8 * 1024 * 1024
SHA256 = re.compile(r"[0-9a-f]{64}\Z")
SHA1 = re.compile(r"[0-9a-f]{40}\Z")
VERSION = re.compile(r"(?:0|[1-9][0-9]*)\.(?:0|[1-9][0-9]*)\.(?:0|[1-9][0-9]*)\Z")
# These are the existing CI package-input rules, retained for the Git-free kit.
WINDOWS_RESERVED = re.compile(r"(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?\Z", re.IGNORECASE)


class PackageError(ValueError):
    def __init__(self, code: str):
        super().__init__(code)
        self.code = code


def canonical(value: object) -> bytes:
    return (json.dumps(value, sort_keys=True, indent=2, ensure_ascii=True) + "\n").encode("ascii")


def safe_relative(path: str) -> bool:
    if not isinstance(path, str) or not path or "\\" in path:
        return False
    parts = path.split("/")
    return not PurePosixPath(path).is_absolute() and all(
        part not in {"", ".", ".."} and part[-1] not in {".", " "}
        and not WINDOWS_RESERVED.fullmatch(part)
        and not any(ord(char) < 32 or ord(char) == 127 or char in '<>:"|?*' for char in part)
        for part in parts)


def ordinary(path: Path, directory: bool = False) -> None:
    info = path.lstat()
    if stat.S_ISLNK(info.st_mode) or getattr(info, "st_file_attributes", 0) & 0x400 \
            or not (stat.S_ISDIR(info.st_mode) if directory else stat.S_ISREG(info.st_mode)):
        raise PackageError("ordinary_path_required")


def read_bounded(path: Path, maximum: int) -> bytes:
    ordinary(path)
    if not 0 < path.stat().st_size <= maximum:
        raise PackageError("invalid_asset_size")
    with path.open("rb") as stream:
        data = stream.read(maximum + 1)
    if not 0 < len(data) <= maximum:
        raise PackageError("invalid_asset_size")
    return data


def archive_bytes(payload: dict[str, bytes]) -> bytes:
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w", compression=zipfile.ZIP_STORED) as archive:
        for path in sorted(payload):
            info = zipfile.ZipInfo(path, date_time=TIMESTAMP)
            info.create_system = 3
            info.external_attr = FILE_ATTRIBUTES
            archive.writestr(info, payload[path])
    return stream.getvalue()


def read_archive(archive: bytes, sizes: dict[str, int]) -> dict[str, bytes]:
    """Inspect every bounded header before reading any member's payload."""
    if not 0 < len(archive) <= MAX_ARCHIVE_BYTES:
        raise PackageError("invalid_asset_size")
    try:
        # Bound the central directory before ZipFile allocates member objects.
        end = archive.rfind(b"PK\x05\x06")
        if end < 0 or end + 22 > len(archive):
            raise PackageError("invalid_package")
        _, disk, start_disk, disk_count, count, central_size, central_offset, comment_size = struct.unpack(
            "<4s4H2LH", archive[end:end + 22])
        if count != len(sizes) or disk_count != count or comment_size:
            raise PackageError("zip_membership_mismatch")
        expected_central_size = sum(46 + len(path.encode("utf-8")) for path in sizes)
        if disk or start_disk or central_size != expected_central_size:
            raise PackageError("invalid_zip_member")
        # ZIP64 can override the bounded EOCD; reject it before archive parsing.
        if end >= 20 and archive[end - 20:end - 16] == b"PK\x06\x07" \
                or end + 22 != len(archive) or central_offset + central_size != end:
            raise PackageError("zip_recipe_mismatch")
        with zipfile.ZipFile(io.BytesIO(archive), "r") as source:
            members = source.infolist()
            if len(members) != len(sizes) or [item.filename for item in members] != sorted(sizes) \
                    or source.comment:
                raise PackageError("zip_membership_mismatch")
            for member in members:
                path = member.filename
                if not safe_relative(path) or member.orig_filename != path \
                        or member.date_time != TIMESTAMP or member.compress_type != zipfile.ZIP_STORED \
                        or member.flag_bits != 0 or member.create_system != 3 \
                        or member.external_attr != FILE_ATTRIBUTES or member.internal_attr != 0 \
                        or member.extra or member.comment or member.is_dir() \
                        or member.file_size != sizes[path] or member.compress_size != member.file_size:
                    raise PackageError("invalid_zip_member")
            return {member.filename: source.read(member) for member in members}
    except (KeyError, TypeError, zipfile.BadZipFile, RuntimeError, NotImplementedError, struct.error) as error:
        raise PackageError("invalid_package") from error


def validate_archive(archive: bytes, payload: dict[str, bytes]) -> None:
    """The existing exact recipe, shared by Git and frozen-byte verification."""
    if read_archive(archive, {path: len(data) for path, data in payload.items()}) != payload:
        raise PackageError("zip_payload_mismatch")
    if archive != archive_bytes(payload):
        raise PackageError("zip_recipe_mismatch")


def _strict_json(data: bytes, *, canonical_required: bool = True) -> dict:
    def pairs(items):
        value = {}
        for key, item in items:
            if key in value:
                raise PackageError("invalid_package")
            value[key] = item
        return value

    def constant(_):
        raise PackageError("invalid_package")

    try:
        value = json.loads(data.decode("utf-8", errors="strict"), object_pairs_hook=pairs, parse_constant=constant)
    except (UnicodeError, json.JSONDecodeError, RecursionError) as error:
        raise PackageError("invalid_package") from error
    if type(value) is not dict or canonical_required and data != canonical(value):
        raise PackageError("manifest_mismatch")
    return value


def dependency_metadata(policy_bytes: bytes, requirements_bytes: bytes) -> dict:
    """Project typed, externally bound policy metadata; no runtime acceptance.

    The developer producer separately reuses the shipped strict validator.
    The frozen consumer checks this closed metadata contract without importing
    or executing any application payload helper.
    """
    if type(policy_bytes) is not bytes or not 0 < len(policy_bytes) <= 32768 \
            or type(requirements_bytes) is not bytes or not 0 < len(requirements_bytes) <= 4096:
        raise PackageError("invalid_dependency_policy")
    try:
        policy = _strict_json(policy_bytes, canonical_required=False)
        expected = {"schema_version", "policy_version", "reviewed_on", "python", "pypdf", "calibre",
                    "excluded_python_branches", "sources"}
        if set(policy) != expected or type(policy["schema_version"]) is not int or policy["schema_version"] != 1 \
                or type(policy["policy_version"]) is not int or policy["policy_version"] < 1:
            raise PackageError("invalid_dependency_policy")
        python = policy["python"]
        if type(python) is not dict or set(python) != {"implementation", "platform", "machine", "bits", "stable_only", "gil_disabled", "branches"} \
                or python["implementation"] != "cpython" or python["platform"] != "win32" or python["machine"] != "AMD64" \
                or type(python["bits"]) is not int or python["bits"] != 64 or python["stable_only"] is not True \
                or python["gil_disabled"] is not False or type(python["branches"]) is not list or not python["branches"]:
            raise PackageError("invalid_dependency_policy")
        branch_keys = {"major", "minor", "compatibility_min_patch", "security_min_patch", "recommended_patch", "blocked_patches", "security_sources"}
        seen = set()
        for branch in python["branches"]:
            if type(branch) is not dict or set(branch) != branch_keys \
                    or any(type(branch[key]) is not int or not 0 <= branch[key] <= 1000000 for key in
                           ("major", "minor", "compatibility_min_patch", "security_min_patch", "recommended_patch")) \
                    or branch["major"] != 3 or (branch["major"], branch["minor"]) in seen \
                    or branch["recommended_patch"] < max(branch["compatibility_min_patch"], branch["security_min_patch"]) \
                    or type(branch["blocked_patches"]) is not list \
                    or any(type(value) is not int or not 0 <= value <= 1000000 for value in branch["blocked_patches"]) \
                    or len(branch["blocked_patches"]) != len(set(branch["blocked_patches"])) \
                    or type(branch["security_sources"]) is not list or not branch["security_sources"] \
                    or any(type(value) is not str or not re.fullmatch(r"[A-Za-z][A-Za-z0-9_-]{0,63}", value) for value in branch["security_sources"]):
                raise PackageError("invalid_dependency_policy")
            seen.add((branch["major"], branch["minor"]))
        excluded = policy["excluded_python_branches"]
        if type(excluded) is not list or any(type(branch) is not dict or set(branch) != {"major", "minor", "reason"} \
                or any(type(branch[key]) is not int or not 0 <= branch[key] <= 1000000 for key in ("major", "minor")) \
                or type(branch["reason"]) is not str or not 0 < len(branch["reason"]) <= 1024 for branch in excluded):
            raise PackageError("invalid_dependency_policy")
        pypdf, calibre = policy["pypdf"], policy["calibre"]
        if type(pypdf) is not dict or set(pypdf) != {"version", "wheel_sha256"} \
                or type(pypdf["version"]) is not str or not VERSION.fullmatch(pypdf["version"]) \
                or type(pypdf["wheel_sha256"]) is not str or not SHA256.fullmatch(pypdf["wheel_sha256"]) \
                or type(calibre) is not dict or set(calibre) != {"version"} \
                or type(calibre["version"]) is not str or not VERSION.fullmatch(calibre["version"]):
            raise PackageError("invalid_dependency_policy")
        requirement = "pypdf==" + pypdf["version"] + " --hash=sha256:" + pypdf["wheel_sha256"]
        lines = [line.strip() for line in requirements_bytes.decode("utf-8-sig", errors="strict").splitlines()
                 if line.strip() and not line.strip().startswith("#")]
        if lines != [requirement]:
            raise PackageError("invalid_dependency_policy")
        return {
            "windows": {"supported": ["Windows 10 22H2 x64", "Windows 11 x64"], "bundled": False},
            "powershell": {"supported": ["Windows PowerShell 5.1", "PowerShell 7.6.6"], "bundled": False},
            "python": {"architecture": "x64", "flavor": "regular GIL CPython", "bundled": False,
                       "policy_path": "engine/WinBookSplit.Compatibility.json", "policy_sha256": sha256(policy_bytes).hexdigest(),
                       "policy_schema_version": policy["schema_version"], "policy_version": policy["policy_version"],
                       **python, "excluded_branches": excluded},
            "pypdf": {"version": pypdf["version"], "wheel_sha256": pypdf["wheel_sha256"], "requirement": requirement, "bundled": False},
            "calibre": {"version": calibre["version"], "required_for": ["epub", "azw3"], "bundled": False}}
    except (KeyError, TypeError, ValueError, UnicodeError, RecursionError) as error:
        raise PackageError("invalid_dependency_policy") from error


def _bindings(expectations: dict) -> tuple[str, tuple[str, str, str]]:
    expected_keys = {"version", "source_commit", "source_tree", "payload_sha256", "policy_sha256", "assets"}
    if type(expectations) is not dict or set(expectations) != expected_keys:
        raise PackageError("invalid_asset_binding")
    version = expectations["version"]
    if type(version) is not str or len(version) > 32 or not VERSION.fullmatch(version):
        raise PackageError("invalid_asset_binding")
    for name in ("source_commit", "source_tree"):
        if type(expectations[name]) is not str or not SHA1.fullmatch(expectations[name]):
            raise PackageError("invalid_asset_binding")
    for name in ("payload_sha256", "policy_sha256"):
        if type(expectations[name]) is not str or not SHA256.fullmatch(expectations[name]):
            raise PackageError("invalid_asset_binding")
    names = ("WinBookSplit-v" + version + ".zip", "release-manifest.json", "SHA256SUMS.txt")
    assets = expectations["assets"]
    if type(assets) is not dict or set(assets) != set(names):
        raise PackageError("invalid_asset_binding")
    limits = (MAX_ARCHIVE_BYTES, MAX_MANIFEST_BYTES, 512)
    for name, limit in zip(names, limits):
        binding = assets[name]
        if type(binding) is not dict or set(binding) != {"bytes", "sha256"} \
                or type(binding["bytes"]) is not int or not 0 < binding["bytes"] <= limit \
                or type(binding["sha256"]) is not str or not SHA256.fullmatch(binding["sha256"]):
            raise PackageError("invalid_asset_binding")
    return version, names


def _frozen_payload(archive: bytes, manifest: dict, expectations: dict) -> dict[str, bytes]:
    manifest_keys = {"schema_version", "application", "version", "source_commit", "source_tree", "payload_sha256",
                     "files", "runtime_sibling_dependencies", "dependencies", "builder", "signed", "scope", "archive"}
    if set(manifest) != manifest_keys or type(manifest["schema_version"]) is not int or manifest["schema_version"] != 1 \
            or manifest["application"] != "WinBookSplit" or manifest["signed"] is not False \
            or any(manifest[name] != expectations[name] for name in ("version", "source_commit", "source_tree", "payload_sha256")):
        raise PackageError("manifest_mismatch")
    rows = manifest["files"]
    if type(rows) is not list or not rows:
        raise PackageError("manifest_mismatch")
    paths = []
    total = 0
    for row in rows:
        if type(row) is not dict or set(row) != {"path", "scope", "git_blob", "git_mode", "bytes", "sha256"} \
                or not safe_relative(row["path"]) or type(row["bytes"]) is not int \
                or type(row["scope"]) is not str or row["scope"] not in {"runtime", "support_source"} \
                or not 0 < row["bytes"] <= MAX_FILE_BYTES or type(row["git_mode"]) is not str \
                or row["git_mode"] not in {"100644", "100755"} \
                or type(row["git_blob"]) is not str or not SHA1.fullmatch(row["git_blob"]) \
                or type(row["sha256"]) is not str or not SHA256.fullmatch(row["sha256"]):
            raise PackageError("manifest_mismatch")
        paths.append(row["path"])
        total += row["bytes"]
    if paths != sorted(paths) or len(paths) != len(set(path.casefold() for path in paths)) \
            or total > MAX_ARCHIVE_BYTES or sha256(canonical(rows)).hexdigest() != manifest["payload_sha256"]:
        raise PackageError("manifest_mismatch")
    file_names = {path.casefold() for path in paths}
    directories = {}
    for path in paths:
        parts = path.split("/")
        for length in range(1, len(parts)):
            directory = "/".join(parts[:length])
            folded = directory.casefold()
            if folded in file_names or folded in directories and directories[folded] != directory:
                raise PackageError("manifest_mismatch")
            directories[folded] = directory
    expected_archive = {"name": "WinBookSplit-v" + expectations["version"] + ".zip",
                        "bytes": len(archive), "sha256": sha256(archive).hexdigest()}
    if manifest["archive"] != expected_archive:
        raise PackageError("manifest_mismatch")
    sizes = {row["path"]: row["bytes"] for row in rows}
    payload = read_archive(archive, sizes)
    for row in rows:
        data = payload[row["path"]]
        if len(data) != row["bytes"] or sha256(data).hexdigest() != row["sha256"] \
                or sha1(b"blob " + str(len(data)).encode("ascii") + b"\0" + data).hexdigest() != row["git_blob"]:
            raise PackageError("zip_payload_mismatch")
    if archive != archive_bytes(payload):
        raise PackageError("zip_recipe_mismatch")
    if "VERSION" not in payload or payload["VERSION"].decode("utf-8-sig", errors="strict").strip() != expectations["version"] \
            or "engine/WinBookSplit.Compatibility.json" not in payload \
            or sha256(payload["engine/WinBookSplit.Compatibility.json"]).hexdigest() != expectations["policy_sha256"]:
        raise PackageError("manifest_mismatch")
    if "requirements.txt" not in payload or canonical(manifest["dependencies"]) != canonical(dependency_metadata(
            payload["engine/WinBookSplit.Compatibility.json"], payload["requirements.txt"])):
        raise PackageError("manifest_mismatch")
    return payload


def verify_frozen_assets(assets_root: Path, expectations: dict) -> tuple[dict, dict[str, bytes]]:
    """Return fully verified in-memory bytes for the caller's guarded extraction.

    The caller must authenticate this helper and the external expectations before
    invocation. This function never extracts, launches or removes a file. Its
    source commit/tree claim depends on the prior independent Git-backed build.
    """
    _, names = _bindings(expectations)
    root = Path(assets_root)
    if not root.is_absolute() or root.drive.startswith("\\\\") or any(part in {".", ".."} for part in root.parts):
        raise PackageError("absolute_output_required")
    for parent in reversed(root.parents):
        ordinary(parent, directory=True)
    ordinary(root, directory=True)
    if root.resolve(strict=True) != root or {path.name for path in root.iterdir()} != set(names):
        raise PackageError("asset_membership_mismatch")
    raw = {}
    for name, maximum in zip(names, (MAX_ARCHIVE_BYTES, MAX_MANIFEST_BYTES, 512)):
        data = read_bounded(root / name, maximum)
        expected = expectations["assets"][name]
        if len(data) != expected["bytes"] or sha256(data).hexdigest() != expected["sha256"]:
            raise PackageError("asset_hash_mismatch")
        raw[name] = data
    manifest = _strict_json(raw[names[1]])
    sums = (sha256(raw[names[0]]).hexdigest() + "  " + names[0] + "\n" +
            sha256(raw[names[1]]).hexdigest() + "  " + names[1] + "\n").encode("ascii")
    if raw[names[2]] != sums:
        raise PackageError("checksum_mismatch")
    return manifest, _frozen_payload(raw[names[0]], manifest, expectations)
