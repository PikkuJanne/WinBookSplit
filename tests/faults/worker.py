"""Fixed AC-073 child: real engine writes plus one declared second-write fault.

This development-only interposer never changes shipped source. It launches no
grandchildren, and accepts only two barrier roles and three ordinary run IDs.
"""

import argparse
from hashlib import sha256
import importlib.util
import json
import os
from pathlib import Path
import sys
import time
from unittest.mock import patch

import pypdf


FIXED_TIMESTAMP = "20261010-120000"
PARTIAL_BYTES = b"WBS authored incomplete second PDF for AC-073\n"
ROLES = {"prior-repeat-1", "prior-repeat-2", "overlap-success", "overlap-mid-write", "repeat-after"}


def emit(path, record):
    pending = path.with_name(path.name + ".pending")
    with pending.open("x", encoding="utf-8", newline="\n") as stream:
        json.dump(record, stream, ensure_ascii=True, sort_keys=True)
        stream.write("\n")
        stream.flush()
        os.fsync(stream.fileno())
    os.rename(pending, path)  # Windows refuses an existing destination.


def identity(path):
    details = os.lstat(path)
    return {"device": details.st_dev, "inode": details.st_ino}


def wait_gate(path, nonce):
    deadline = time.monotonic() + 30
    while not path.is_file():
        if time.monotonic() >= deadline:
            raise RuntimeError("Authored AC-073 control deadline elapsed")
        time.sleep(0.01)
    if path.read_bytes() != nonce.encode("ascii"):
        raise RuntimeError("Authored control nonce changed")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--engine", required=True, type=Path)
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--base", required=True, type=Path)
    parser.add_argument("--control", required=True, type=Path)
    parser.add_argument("--role", required=True, choices=sorted(ROLES))
    parser.add_argument("--nonce", required=True)
    args = parser.parse_args()
    if not sys.flags.isolated or not sys.dont_write_bytecode or os.name != "nt":
        raise ValueError("Actual isolated Windows developer Python required")
    if len(args.nonce) != 32 or any(char not in "0123456789abcdef" for char in args.nonce):
        raise ValueError("Authored control nonce required")
    for path in (args.engine, args.source, args.base, args.control):
        if not path.is_absolute() or path.resolve(strict=True) != path:
            raise ValueError("Existing literal absolute paths required")
    spec = importlib.util.spec_from_file_location("wbs_fault_isolation_engine", args.engine)
    engine = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(engine)
    emit(args.control / "started.json", {"protocol": "winbooksplit.test.isolation", "version": 1,
        "nonce": args.nonce, "role": args.role, "phase": "interpreter-held", "pid": os.getpid(),
        "parent_pid": os.getppid(), "python_executable": sys.executable, "python_prefix": sys.prefix,
        "python_base_prefix": sys.base_prefix, "python_version": ".".join(map(str, sys.version_info[:3])),
        "pypdf_version": pypdf.__version__, "pypdf_path": pypdf.__file__})
    wait_gate(args.control / "identity-ack", args.nonce)
    original_slice, original_write = engine.write_slice, engine.PdfWriter.write
    run, slice_count, write_count = None, 0, 0

    def write(writer, stream):
        nonlocal write_count
        write_count += 1
        if args.role != "overlap-mid-write" or write_count != 2:
            return original_write(writer, stream)
        path = Path(stream.name)
        if run is None or path.parent != Path(run.stage) or str(path) not in run.files:
            raise RuntimeError("Second real exclusive stream was not registered by OutputRun")
        written = stream.write(PARTIAL_BYTES)
        stream.flush()
        os.fsync(stream.fileno())
        details = os.fstat(stream.fileno())
        if written != len(PARTIAL_BYTES) or details.st_size != len(PARTIAL_BYTES):
            raise RuntimeError("Authored partial write did not reach actual stream")
        partial = {"protocol": "winbooksplit.test.isolation", "version": 1,
                   "nonce": args.nonce, "role": args.role, "pid": os.getpid(),
                   "phase": "partial-second-write", "run_id": run.run_id,
                   "stage": run.stage, "stage_identity": identity(run.stage),
                   "completed_real_slices": 1, "write_call": write_count,
                   "path": str(path), "device": details.st_dev, "inode": details.st_ino,
                   "size_bytes": details.st_size, "sha256": sha256(PARTIAL_BYTES).hexdigest(),
                   "bytes_written": written,
                   "ledger": {Path(name).name: {"device": item[0], "inode": item[1]}
                              for name, item in run.files.items()}}
        emit(args.control / "partial.json", partial)
        raise OSError("Authored AC-073 failure during second PDF stream write")

    def slice(reader, start, end, path, on_created=None):
        nonlocal run, slice_count
        run = getattr(on_created, "__self__", None)
        if run is None or not isinstance(run, engine.OutputRun):
            raise RuntimeError("Actual OutputRun creation callback required")
        original_slice(reader, start, end, path, on_created=on_created)
        slice_count += 1
        if args.role not in {"overlap-success", "overlap-mid-write"} or slice_count != 1:
            return
        # The original writer has closed the first PDF. Use the production
        # ownership assertion and reopened-PDF validator before holding it.
        run.assert_owned()
        first = engine._validate_output(path, {"start": start, "end": end})
        ready = {"protocol": "winbooksplit.test.isolation", "version": 1,
                 "nonce": args.nonce, "role": args.role, "pid": os.getpid(),
                 "phase": "first-slice-held", "run_id": run.run_id, "stage": run.stage,
                 "stage_identity": identity(run.stage), "completed_real_slices": slice_count,
                 "first_slice": {"filename": Path(path).name, "start": start, "end": end, **first},
                 "ledger": {Path(name).name: {"device": item[0], "inode": item[1]}
                            for name, item in run.files.items()}}
        emit(args.control / "ready.json", ready)
        wait_gate(args.control / "release", args.nonce)
        run.assert_owned()

    engine._timestamp = lambda: FIXED_TIMESTAMP
    sys.argv = [str(args.engine), str(args.source), str(args.base), "manual", "4,7"]
    with patch.object(engine, "write_slice", side_effect=slice), \
            patch.object(engine.PdfWriter, "write", autospec=True, side_effect=write):
        return engine.main()


if __name__ == "__main__":
    raise SystemExit(main())
