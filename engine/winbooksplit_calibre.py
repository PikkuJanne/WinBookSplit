"""Verify and unpack the pinned official Windows Calibre Portable release."""

import argparse
from contextlib import contextmanager
import ctypes
from ctypes import wintypes
from hashlib import sha256, sha512
from http.client import HTTPException
import importlib.util
import json
import lzma
import msvcrt
import os
from pathlib import Path
import re
import shutil
import stat
import struct
import sys
from tempfile import mkdtemp
import time
from urllib.error import URLError
from urllib.parse import urlsplit
from urllib.request import HTTPRedirectHandler, build_opener
from uuid import uuid4
import zipfile
from zlib import crc32


VERSION = "9.18.0"
ARTIFACT = f"calibre-portable-installer-{VERSION}.exe"
URL = f"https://download.calibre-ebook.com/{VERSION}/{ARTIFACT}"
SIZE = 259196168
SHA256 = "bb3b5423e1d64a735322a006edce15c6fb012980c2c5e55c1c45eaeba21a28dc"
SHA512 = ("6ed05950375b098b58ea2e71cde927fe88623074f5f703246f5cd36d8c91c79625"
          "ba5a184134b3e3ff005843491af3d1ee2a1424d8a3e2e2bee552465d559dfe")
TREE_SHA256 = "5de8584ed4aaa7516cbac1aae102ca6f40f7af9a16eb351b100de8fce418a3a4"
CURRENT = "current.json"
CONVERTER = Path("Calibre", "ebook-convert.exe")
CHUNK = 1024 * 1024
MAX_PAYLOAD = 2 * 1024**3
MAX_MEMBERS = 20000


class CalibreError(Exception):
    def __init__(self, code, message):
        super().__init__(message)
        self.code = code


def ordinary(path, *, directory=False):
    if not path.is_absolute() or not re.fullmatch(r"[A-Za-z]:", path.drive) or ".." in path.parts:
        raise OSError("Managed Calibre requires an ordinary absolute local path.")
    info = path.lstat()
    if not (stat.S_ISDIR(info.st_mode) if directory else stat.S_ISREG(info.st_mode) and info.st_nlink == 1):
        raise OSError("Managed Calibre path has an unexpected file type.")
    for component in (path, *path.parents):
        if component.lstat().st_file_attributes & stat.FILE_ATTRIBUTE_REPARSE_POINT:
            raise OSError("Managed Calibre paths cannot use reparse points.")
    return info


def file_hash(path):
    with path.open("rb") as stream:
        return sha256_stream(stream)


def sha256_stream(stream):
    digest = sha256()
    for block in iter(lambda: stream.read(CHUNK), b""):
        digest.update(block)
    return digest.hexdigest()


def artifact_valid(path):
    if ordinary(path).st_size != SIZE:
        raise CalibreError("converter_integrity_failed", "The Calibre Portable installer has an unexpected size.")
    sha256_digest, sha512_digest = sha256(), sha512()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(CHUNK), b""):
            sha256_digest.update(block)
            sha512_digest.update(block)
    if sha256_digest.hexdigest() != SHA256 or sha512_digest.hexdigest() != SHA512:
        raise CalibreError("converter_integrity_failed", "The Calibre Portable installer failed the pinned integrity checks.")


class OfficialRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        url = urlsplit(newurl)
        if url.scheme != "https" or url.netloc != "download.calibre-ebook.com":
            raise URLError("Unexpected Calibre download redirect.")
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def download(path):
    started = time.monotonic()
    try:
        with build_opener(OfficialRedirect).open(URL, timeout=20) as response, path.open("xb") as stream:
            size = 0
            for block in iter(lambda: response.read(CHUNK), b""):
                size += len(block)
                if time.monotonic() - started > 300:
                    raise TimeoutError("The Calibre download exceeded its time limit.")
                if size > SIZE:
                    raise CalibreError("converter_integrity_failed", "The Calibre download exceeds its pinned size.")
                stream.write(block)
            if size != SIZE:
                raise CalibreError("converter_download_failed", "The official Calibre download ended before the expected size.")
    except (OSError, HTTPException) as error:
        raise CalibreError("converter_download_failed", f"Cannot download the official Calibre Portable package: {error}") from error


@contextmanager
def portable_resource(installer):
    """Map the PE file as data so no installer code or entry point runs."""
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.LoadLibraryExW.argtypes = (wintypes.LPCWSTR, wintypes.HANDLE, wintypes.DWORD)
    kernel.LoadLibraryExW.restype = wintypes.HMODULE
    kernel.FindResourceW.argtypes = (wintypes.HMODULE, wintypes.LPCWSTR, wintypes.LPCWSTR)
    kernel.FindResourceW.restype = wintypes.HANDLE
    kernel.SizeofResource.argtypes = (wintypes.HMODULE, wintypes.HANDLE)
    kernel.SizeofResource.restype = wintypes.DWORD
    kernel.LoadResource.argtypes = (wintypes.HMODULE, wintypes.HANDLE)
    kernel.LoadResource.restype = wintypes.HANDLE
    kernel.LockResource.argtypes = (wintypes.HANDLE,)
    kernel.LockResource.restype = ctypes.c_void_p
    kernel.FreeLibrary.argtypes = (wintypes.HMODULE,)
    kernel.FreeLibrary.restype = wintypes.BOOL
    handle = kernel.LoadLibraryExW(str(installer), None, 0x40 | 0x20)
    if not handle:
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        resource = kernel.FindResourceW(handle, "extra", "extra")
        size = kernel.SizeofResource(handle, resource) if resource else 0
        memory = kernel.LoadResource(handle, resource) if size else None
        address = kernel.LockResource(memory) if memory else None
        if not address or not 26 < size < SIZE:
            raise OSError("The verified installer has no usable portable archive resource.")
        yield address, size
    finally:
        kernel.FreeLibrary(handle)


def decompress_portable(installer, destination):
    with portable_resource(installer) as (address, size), destination.open("xb") as stream:
        header = ctypes.string_at(address, 6)
        code = header[-1]
        dictionary = 1 << (code & 31)
        dictionary -= dictionary // 16 * (code >> 5)
        # This signed release uses the legacy LZIP v0 trailer, without a
        # member-length field. The PE resource supplies the fixed input size.
        expected_crc, expected_size = struct.unpack("<IQ", ctypes.string_at(address + size - 12, 12))
        if header[:5] != b"LZIP\x00" or not 4096 <= dictionary <= 128 * CHUNK \
                or not 0 < expected_size <= MAX_PAYLOAD:
            raise ValueError("The portable archive has an unsupported LZIP header or trailer.")
        decoder = lzma.LZMADecompressor(format=lzma.FORMAT_RAW, filters=[
            {"id": lzma.FILTER_LZMA1, "dict_size": dictionary, "lc": 3, "lp": 0, "pb": 2},
        ])
        offset, end, count, checksum = 6, size - 12, 0, 0
        while not decoder.eof:
            data = b""
            if decoder.needs_input:
                length = min(CHUNK, end - offset)
                if not length:
                    raise ValueError("The portable archive ended before its LZIP end marker.")
                data = ctypes.string_at(address + offset, length)
                offset += length
            output = decoder.decompress(data, max_length=CHUNK)
            count += len(output)
            if count > expected_size:
                raise ValueError("The portable archive exceeds its declared expansion size.")
            checksum = crc32(output, checksum)
            stream.write(output)
        if decoder.unused_data or offset != end or count != expected_size or checksum != expected_crc:
            raise ValueError("The portable archive failed LZIP length or CRC validation.")


def extract_portable(archive, destination):
    destination.mkdir()
    entries, total = set(), 0
    with zipfile.ZipFile(archive) as source:
        if not 0 < len(source.infolist()) <= MAX_MEMBERS:
            raise ValueError("The portable ZIP has an unsupported member count.")
        for entry in source.infolist():
            parts = entry.filename.removesuffix("/").split("/")
            alias = "/".join(parts).casefold()
            if alias in entries or any(not part or part in {".", ".."} or part[-1] in " ." or
                    os.path.isreserved(part) or re.search(r'[\\:<>|"?*\x00-\x1f]', part) for part in parts):
                raise ValueError("The portable ZIP contains an unsafe or ambiguous member path.")
            entries.add(alias)
            mode = stat.S_IFMT(entry.external_attr >> 16)
            if mode not in {0, stat.S_IFDIR if entry.is_dir() else stat.S_IFREG}:
                raise ValueError("The portable ZIP contains an unsupported member type.")
            target = destination.joinpath(*parts)
            if entry.is_dir():
                target.mkdir(parents=True, exist_ok=True)
                continue
            total += entry.file_size
            if total > MAX_PAYLOAD:
                raise ValueError("The portable ZIP exceeds its expansion limit.")
            target.parent.mkdir(parents=True, exist_ok=True)
            with source.open(entry) as data, target.open("xb") as output:
                shutil.copyfileobj(data, output, CHUNK)
            if target.stat().st_size != entry.file_size:
                raise ValueError("The portable ZIP member length differs after extraction.")


def tree_digest(root):
    result = sha256()
    members = []
    count, total = 0, 0
    for directory, subdirs, files in os.walk(root, followlinks=False):
        parent = Path(directory)
        for name in subdirs:
            child = parent / name
            ordinary(child, directory=True)
            members.append((child.relative_to(root).as_posix() + "/", None))
        count += len(subdirs) + len(files)
        if count > MAX_MEMBERS:
            raise ValueError("The managed Calibre tree exceeds its supported member count.")
        for name in files:
            path = parent / name
            total += ordinary(path).st_size
            if total > MAX_PAYLOAD:
                raise ValueError("The managed Calibre tree exceeds its supported size.")
            members.append((path.relative_to(root).as_posix(), path))
    for name, path in sorted(members):
        result.update(name.encode("utf-8") + b"\x00")
        if path is not None:
            result.update(bytes.fromhex(file_hash(path)))
    return result.hexdigest()


def current(root):
    pointer = root / CURRENT
    if not os.path.lexists(pointer):
        return None
    try:
        if ordinary(pointer).st_size > 512:
            raise ValueError("Oversized managed Calibre record.")
        data = pointer.read_bytes()
        record = json.loads(data)
        name = record.get("directory") if isinstance(record, dict) else None
        if not isinstance(name, str) or not re.fullmatch(re.escape(VERSION) + r"-[0-9a-f]{32}", name) or record != {
                "directory": name, "version": VERSION, "installer_sha256": SHA256, "tree_sha256": TREE_SHA256} \
                or data != json.dumps(record, sort_keys=True).encode("ascii"):
            raise ValueError("Unexpected managed Calibre identity.")
        directory = root / name
        ordinary(directory, directory=True)
        if tree_digest(directory) != TREE_SHA256:
            raise ValueError("Managed Calibre files differ from the reviewed portable distribution.")
        executable = directory / CONVERTER
        ordinary(executable)
        return executable
    except (OSError, ValueError) as error:
        raise CalibreError("converter_managed_invalid", str(error)) from error


@contextmanager
def preparation_lock(root):
    lock = root / ".prepare.lock"
    if os.path.lexists(lock):
        ordinary(lock)
    with lock.open("a+b") as handle:
        handle.seek(0)
        try:
            msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError as error:
            raise CalibreError("converter_setup_failed", "Another WinBookSplit Calibre preparation is running.") from error
        try:
            yield
        finally:
            handle.seek(0)
            msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)


@contextmanager
def owned_scratch(root):
    path = Path(__file__).resolve().with_name("winbooksplit_windows.py")
    spec = importlib.util.spec_from_file_location("_wbs_calibre_windows", path)
    windows = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(windows)
    with windows.DirectoryGuard(root.parent), windows.DirectoryGuard(root) as base:
        scratch = Path(mkdtemp(prefix=".prepare-", dir=root))
        with windows.DirectoryGuard(scratch, delete_access=True) as stage:
            base.allow_publication(stage)
            try:
                yield scratch, stage, windows
            finally:
                stage.assert_unchanged()
                children = list(scratch.iterdir())
                if any(child.name not in {ARTIFACT, "portable.zip", "portable", CURRENT} for child in children):
                    raise OSError("The owned Calibre preparation folder contains an unexpected member.")
                for child in children:
                    if child.name == "portable":
                        ordinary(child, directory=True)
                        # Only this private extraction subtree can be removed.
                        # Windows shutil does not traverse junctions or symlinks.
                        shutil.rmtree(child)
                    else:
                        ordinary(child)
                        with windows.FileGuard(child, delete_access=True) as file:
                            file.delete()
                base.restrict_writes(stage)
                stage.delete()


def prepare(root, offline):
    base = root.parent.parent
    ordinary(base, directory=True)
    for directory in (root.parent, root):
        directory.mkdir(exist_ok=True)
        ordinary(directory, directory=True)
    with preparation_lock(root):
        try:
            executable = current(root)
        except CalibreError:
            # Keep the previous distribution for inspection. A fresh verified
            # generation can supersede its ordinary pointer atomically.
            ordinary(root / CURRENT)
            executable = None
        if executable is not None:
            return executable, False
        with owned_scratch(root) as (scratch, stage, windows):
            installer = scratch / ARTIFACT
            if offline is None:
                download(installer)
            else:
                if ordinary(offline).st_size != SIZE:
                    raise CalibreError("converter_integrity_failed", "The offline Calibre installer has an unexpected size.")
                with offline.open("rb") as source, installer.open("xb") as output:
                    remaining = SIZE
                    while remaining:
                        block = source.read(min(remaining, CHUNK))
                        if not block:
                            break
                        output.write(block)
                        remaining -= len(block)
                    if remaining or source.read(1):
                        raise CalibreError("converter_integrity_failed", "The offline installer changed while being read.")
            archive, extracted = scratch / "portable.zip", scratch / "portable"
            with windows.FileGuard(installer):
                artifact_valid(installer)
                decompress_portable(installer, archive)
            extract_portable(archive, extracted)
            unpacked = extracted / "Calibre Portable"
            executable = unpacked / CONVERTER
            ordinary(executable)
            if tree_digest(unpacked) != TREE_SHA256:
                raise CalibreError("converter_integrity_failed", "The extracted Calibre tree differs from its reviewed identity.")
            name = VERSION + "-" + uuid4().hex
            target = root / name
            stage.assert_unchanged()
            unpacked.rename(target)
            pointer = scratch / CURRENT
            pointer.write_text(json.dumps({"directory": name, "version": VERSION, "installer_sha256": SHA256,
                                           "tree_sha256": TREE_SHA256}, sort_keys=True), encoding="ascii")
            pointer.replace(root / CURRENT)
            return target / CONVERTER, True


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("action", choices=("find", "prepare"))
    parser.add_argument("root", type=Path)
    parser.add_argument("--offline", type=Path)
    args = parser.parse_args()
    changed = False
    try:
        if args.root.name != "Calibre" or args.root.parent.name != "WinBookSplit" or not args.root.is_absolute():
            raise ValueError("Invalid managed Calibre location.")
        if args.action == "find":
            executable = current(args.root)
        else:
            executable, changed = prepare(args.root, args.offline)
        record = {"code": "ready" if executable else "absent", "path": str(executable) if executable else None,
                  "version": VERSION if executable else None, "changed": changed}
    except CalibreError as error:
        record = {"code": error.code, "message": str(error)}
    except (OSError, ValueError, lzma.LZMAError, zipfile.BadZipFile) as error:
        record = {"code": "converter_setup_failed" if args.action == "prepare" else "converter_managed_invalid",
                  "message": str(error)}
    print(json.dumps(record, ensure_ascii=True), flush=True)
    return 0 if record["code"] in {"ready", "absent"} else 3


if __name__ == "__main__":
    sys.exit(main())
