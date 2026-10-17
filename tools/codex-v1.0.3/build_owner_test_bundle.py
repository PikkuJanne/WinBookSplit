#!/usr/bin/env python3
"""Package the unchanged frozen native kit with simple owner launch instructions.

Developer tool only: the target computer needs neither Python nor Git to start
the delivered test. Delivery assets are read from raw committed Git blobs. This
does not rebuild the native kit, production ZIP, or any execution receipt.
"""
from __future__ import annotations

import argparse
from contextlib import ExitStack
from hashlib import sha256
from io import BytesIO
import json
import os
from pathlib import Path, PurePosixPath
import re
import stat
import subprocess
import sys
from types import ModuleType
import zipfile


SOURCE_R = "243d31ee03efac35dc6f7e88c0c1dfd1203081ed"
SOURCE_TREE = "ea0db1cc5f5494b508f140deaf25eff87a7e39d9"
KIT_NAME = "WinBookSplit-v103-native-kit-243d31e.zip"
KIT_BYTES = 1559722
KIT_SHA256 = "e2f2a36a727edec378b891f5b125a18bea89dc8e41946719ac8b8a6d2e1837fc"
BINDINGS_SHA256 = "446ef20b5aad8462f413ab6c43df12baa8c6d46211e921c4933bea3c3e3e818b"
PRODUCTION_NAME = "assets/WinBookSplit-v1.0.3.zip"
PRODUCTION_BYTES = 932894
PRODUCTION_SHA256 = "70a76bd04f5645e195972bf075ff53d651c9993696a40065241390805d2ee1de"
BUNDLE_NAME = "WinBookSplit-Windows10-Test-243d31e.zip"
DELIVERY_PREFIX = "tools/codex-v1.0.3/owner-test-delivery/"
DELIVERY_FILES = ("START-HERE.html", "READ-ME.txt", "START-TEST.cmd", "Start-OwnerTest.ps1")
MAX_ASSET_BYTES = 1024 * 1024
MAX_KIT_MEMBER_BYTES = 4 * 1024 * 1024
MAX_KIT_EXPANDED_BYTES = 16 * 1024 * 1024
MAX_BUNDLE_BYTES = 24 * 1024 * 1024
SHA1 = re.compile(r"[0-9a-f]{40}\Z")
WINDOWS_RESERVED = re.compile(r"(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?\Z", re.IGNORECASE)
REPOSITORY_ENV_KEYS = frozenset({"GIT_DIR", "GIT_WORK_TREE", "GIT_INDEX_FILE", "GIT_OBJECT_DIRECTORY",
    "GIT_ALTERNATE_OBJECT_DIRECTORIES", "GIT_COMMON_DIR", "GIT_NAMESPACE"})


class BundleError(ValueError):
    pass


def need(condition: bool, message: str) -> None:
    if not condition:
        raise BundleError(message)


def canonical(value: object) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=True) + "\n").encode("ascii")


def identity(data: bytes) -> dict:
    return {"bytes": len(data), "sha256": sha256(data).hexdigest()}


def file_rows(payload: dict[str, bytes]) -> dict:
    return {name: identity(data) for name, data in sorted(payload.items())}


def safe_relative(name: str) -> bool:
    if not isinstance(name, str) or not name or "\\" in name or PurePosixPath(name).is_absolute():
        return False
    return all(part not in {"", ".", ".."} and part[-1] not in {".", " "}
        and not WINDOWS_RESERVED.fullmatch(part)
        and not any(ord(char) < 32 or ord(char) == 127 or char in '<>:"|?*' for char in part)
        for part in name.split("/"))


def ordinary(path: Path, *, directory: bool = False) -> os.stat_result:
    info = path.lstat()
    need(not stat.S_ISLNK(info.st_mode)
         and not (getattr(info, "st_file_attributes", 0) & stat.FILE_ATTRIBUTE_REPARSE_POINT),
         "Refusing a linked or reparse path.")
    need(stat.S_ISDIR(info.st_mode) if directory else stat.S_ISREG(info.st_mode),
         "Expected an ordinary directory." if directory else "Expected an ordinary file.")
    return info


def path_chain(path: Path) -> list[tuple[Path, tuple[int, int]]]:
    chain = []
    for item in (*reversed(path.parents), path):
        info = ordinary(item, directory=True)
        chain.append((item, (info.st_dev, info.st_ino)))
    return chain


def assert_chain(chain: list[tuple[Path, tuple[int, int]]]) -> None:
    for path, expected in chain:
        current = ordinary(path, directory=True)
        need((current.st_dev, current.st_ino) == expected, "Output directory identity changed.")


def bounded_read(path: Path, maximum: int) -> bytes:
    before = ordinary(path)
    need(0 < before.st_size <= maximum, "File size is outside the delivery bounds.")
    with path.open("rb") as stream:
        opened = os.fstat(stream.fileno())
        need((before.st_dev, before.st_ino) == (opened.st_dev, opened.st_ino), "Input file identity changed.")
        data = stream.read(maximum + 1)
        after = os.fstat(stream.fileno())
    current = ordinary(path)
    need(0 < len(data) <= maximum and len(data) == before.st_size
         and (before.st_dev, before.st_ino, before.st_size, before.st_mtime_ns)
             == (after.st_dev, after.st_ino, after.st_size, after.st_mtime_ns)
         and (before.st_dev, before.st_ino) == (current.st_dev, current.st_ino), "Input file changed while reading.")
    return data


def git(repo: Path, *arguments: str) -> bytes:
    environment = {key: value for key, value in os.environ.items() if key.upper() not in REPOSITORY_ENV_KEYS}
    environment["GIT_NO_REPLACE_OBJECTS"] = "1"
    try:
        result = subprocess.run(["git", "-C", str(repo), *arguments], stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=30, check=False, env=environment)
    except (OSError, subprocess.TimeoutExpired) as error:
        raise BundleError("The bounded Git query did not complete.") from error
    need(result.returncode == 0, "The requested committed delivery source is unavailable.")
    need(len(result.stdout) <= MAX_ASSET_BYTES, "Committed delivery source exceeds its size bound.")
    return result.stdout


def committed_delivery(repo: Path, revision: str) -> tuple[dict[str, bytes], dict]:
    # Resolve once; later reads never follow a moving branch name.
    need(revision and not revision.startswith("-"), "Invalid delivery source revision.")
    commit = git(repo, "rev-parse", "--verify", revision + "^{commit}").decode("ascii").strip()
    need(SHA1.fullmatch(commit) is not None, "Invalid delivery source commit.")
    tree = git(repo, "rev-parse", "--verify", commit + "^{tree}").decode("ascii").strip()
    need(SHA1.fullmatch(tree) is not None, "Invalid delivery source tree.")
    payload = {}
    for name in DELIVERY_FILES:
        source_path = DELIVERY_PREFIX + name
        entry = git(repo, "ls-tree", commit, "--", source_path).decode("ascii").strip()
        fields = entry.split(None, 3)
        need(len(fields) == 4 and fields[0] == "100644" and fields[1] == "blob"
             and fields[3] == source_path, "Delivery source must be an ordinary committed file.")
        size = git(repo, "cat-file", "-s", commit + ":" + source_path).decode("ascii").strip()
        need(size.isdecimal() and 0 < int(size) <= MAX_ASSET_BYTES,
             "Committed delivery source exceeds its size bound.")
        data = git(repo, "cat-file", "blob", commit + ":" + source_path)
        need(len(data) == int(size) and b"\x00" not in data,
             "Delivery source is empty, oversized, or not plain text.")
        data.decode("utf-8-sig", errors="strict")
        payload[name] = data
    return payload, {"commit": commit, "tree": tree}


def frozen_members(raw: bytes) -> dict[str, bytes]:
    need(len(raw) == KIT_BYTES and sha256(raw).hexdigest() == KIT_SHA256,
         "The original native kit does not match the frozen final ZIP.")
    payload = {}
    with zipfile.ZipFile(BytesIO(raw), "r") as archive:
        members = archive.infolist()
        need(len(members) == 35, "The frozen native kit must contain exactly 35 files.")
        folded = set()
        total = 0
        for member in members:
            name = member.filename
            mode = member.external_attr >> 16
            need(safe_relative(name) and not member.is_dir() and not member.flag_bits & 1
                 and (stat.S_IFMT(mode) in {0, stat.S_IFREG})
                 and name.casefold() not in folded, "The frozen native kit contains an unsafe member.")
            need(0 < member.file_size <= MAX_KIT_MEMBER_BYTES, "Frozen kit member exceeds its size bound.")
            folded.add(name.casefold())
            total += member.file_size
            need(total <= MAX_KIT_EXPANDED_BYTES, "Frozen native kit exceeds its expanded size bound.")
            data = archive.read(member)
            need(len(data) == member.file_size, "Frozen native kit member was truncated.")
            payload[name] = data
    need("kit-bindings.json" in payload and sha256(payload["kit-bindings.json"]).hexdigest() == BINDINGS_SHA256,
         "Frozen kit bindings do not match the final identity.")
    need(PRODUCTION_NAME in payload and identity(payload[PRODUCTION_NAME])
         == {"bytes": PRODUCTION_BYTES, "sha256": PRODUCTION_SHA256}, "Frozen production ZIP identity differs.")
    return payload


def archive_bytes(payload: dict[str, bytes]) -> bytes:
    need(len(payload) == 41 and len({name.casefold() for name in payload}) == len(payload),
         "Unexpected transfer bundle file closure.")
    with BytesIO() as stream:
        with zipfile.ZipFile(stream, "w", compression=zipfile.ZIP_STORED, allowZip64=False) as archive:
            for name, data in sorted(payload.items()):
                need(safe_relative(name), "Unsafe transfer bundle member name.")
                member = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                member.create_system = 3
                member.external_attr = (stat.S_IFREG | 0o644) << 16
                archive.writestr(member, data)
        raw = stream.getvalue()
    need(0 < len(raw) <= MAX_BUNDLE_BYTES, "Transfer bundle exceeds its size bound.")
    return raw


def exclusive_write(path: Path, data: bytes, chain: list[tuple[Path, tuple[int, int]]]) -> None:
    assert_chain(chain)
    # O_EXCL prevents adoption or replacement of any existing output, including
    # dangling links. This tool never removes a previous or partially written run.
    descriptor = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL | getattr(os, "O_BINARY", 0), 0o600)
    with os.fdopen(descriptor, "wb") as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())
        written = os.fstat(stream.fileno())
    assert_chain(chain)
    current = ordinary(path)
    need((current.st_dev, current.st_ino) == (written.st_dev, written.st_ino)
         and bounded_read(path, MAX_BUNDLE_BYTES) == data, "New output verification failed; retained for inspection.")


def build(args: argparse.Namespace) -> dict:
    repo = args.repository_root.absolute()
    ordinary(repo, directory=True)
    kit = args.kit.absolute()
    path_chain(kit.parent)
    kit_raw = bounded_read(kit, KIT_BYTES)
    members = frozen_members(kit_raw)
    payload, delivery_source = committed_delivery(repo, args.source_commit)
    payload[KIT_NAME] = kit_raw
    payload.update({"kit/" + name: data for name, data in members.items()})
    manifest = {"protocol": "winbooksplit.owner-test-delivery", "schema_version": 1,
        "delivery_source": delivery_source,
        "frozen_source": {"commit": SOURCE_R, "tree": SOURCE_TREE, "version": "1.0.3"},
        "original_kit": {"path": KIT_NAME, "bytes": KIT_BYTES, "sha256": KIT_SHA256,
                         "bindings_sha256": BINDINGS_SHA256},
        "production_zip": {"path": PRODUCTION_NAME, "bytes": PRODUCTION_BYTES, "sha256": PRODUCTION_SHA256},
        "files": file_rows(payload)}
    payload["delivery-manifest.json"] = canonical(manifest)
    raw = archive_bytes(payload)
    output = args.output.absolute()
    need(output.name == BUNDLE_NAME, "Use the fixed transfer bundle filename.")
    chain = path_chain(output.parent)
    receipt = output.with_suffix(".build.json")
    need(not os.path.lexists(output) and not os.path.lexists(receipt),
         "Output already exists; choose a new empty destination directory.")
    record = {"protocol": "winbooksplit.owner-test-delivery-build", "schema_version": 1,
        "success": True, "delivery_source": delivery_source, "frozen_source": manifest["frozen_source"],
        "original_kit": manifest["original_kit"], "production_zip": manifest["production_zip"],
        "bundle": {"path": BUNDLE_NAME, **identity(raw)}, "files": file_rows(payload),
        "native_acceptance": "NOT_RUN; building a transfer bundle is not native execution evidence."}
    # Reuse the already authenticated frozen Windows path guards. Hold every
    # output ancestor through both exclusive writes; before/after checks alone
    # cannot prevent a concurrent parent-directory replacement.
    need(os.name == "nt", "Transfer bundle publication requires Windows directory guards.")
    windows = ModuleType("winbooksplit_owner_delivery_windows")
    exec(compile(members["engine/winbooksplit_windows.py"],
                 "<authenticated-frozen-Windows-guards>", "exec"), windows.__dict__)
    with ExitStack() as stack:
        guards = [stack.enter_context(windows.DirectoryGuard(path)) for path, _ in chain]
        assert_chain(chain)
        exclusive_write(output, raw, chain)
        output_guard = stack.enter_context(windows.FileGuard(output))
        need(bounded_read(output, MAX_BUNDLE_BYTES) == raw, "Held transfer bundle verification failed.")
        output_guard.assert_unchanged()
        exclusive_write(receipt, canonical(record), chain)
        receipt_guard = stack.enter_context(windows.FileGuard(receipt))
        need(bounded_read(receipt, MAX_BUNDLE_BYTES) == canonical(record), "Held build receipt verification failed.")
        receipt_guard.assert_unchanged()
        for guard in guards:
            guard.assert_unchanged()
    return record


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repository-root", type=Path, default=Path(__file__).absolute().parents[2])
    parser.add_argument("--source-commit", default="HEAD", help="Committed delivery assets; defaults to HEAD.")
    parser.add_argument("--kit", type=Path, required=True, help="The unchanged frozen final native kit ZIP.")
    parser.add_argument("--output", type=Path, required=True,
                        help="New output path ending in " + BUNDLE_NAME + "; parent must already exist.")
    args = parser.parse_args(argv)
    try:
        record = build(args)
        print(json.dumps({"success": True, "bundle": record["bundle"], "native_acceptance": "NOT_RUN"}))
        return 0
    except (BundleError, OSError, UnicodeError, ValueError, KeyError, zipfile.BadZipFile) as error:
        # Deliberately avoid raw OS/Git exception strings containing local paths.
        message = str(error) if isinstance(error, BundleError) else "Transfer bundle build failed; existing outputs were preserved."
        print(json.dumps({"success": False, "error": message}), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
