"""Authored finite subprocess controls; no input-controlled source evaluation."""

import json
import os
from pathlib import Path
import subprocess
import sys
import threading
import time

CAP = 65536
FLOOD = 2 * 1024 * 1024
UNICODE = "Janne ääkköset Größe 日本 中文 한국어"


def frame(message="authored process failure", mode="manual"):
    return {"protocol": "winbooksplit.result", "version": 1, "mode": mode,
            "status": "invalid_input", "code": "invalid_start_pages", "message": message,
            "warnings": [], "fallback_modes": [], "exit_code": 1, "written_count": 0,
            "execution": None}


def encoded_frame(message="authored process failure", mode="manual"):
    return json.dumps(frame(message, mode), ensure_ascii=False, separators=(",", ":")).encode("utf-8")


def emit(fd, data):
    while data:
        count = os.write(fd, data)
        data = data[count:]


def flood(mode="manual"):
    def stderr():
        emit(2, b"E" * FLOOD + b"\nFINAL_STDERR_" + UNICODE.encode("utf-8"))
    thread = threading.Thread(target=stderr)
    thread.start()
    emit(1, b"O" * FLOOD + b"\n" + encoded_frame(mode=mode) + b"\n")
    thread.join()


def write_pid(path):
    Path(path).write_text(json.dumps({"pid": os.getpid()}), encoding="utf-8")


def main():
    kind, *arguments = sys.argv[1:]
    if kind == "flood":
        flood()
        return 1
    if kind == "fast-tail":
        emit(1, b"\n\nstdout blank\n\nno-newline-output")
        emit(2, b"\n\nFINAL_STDERR_NO_NEWLINE")
        return 23
    if kind == "unicode-boundary":
        emit(1, ("日" * (CAP // 3 + 1) + UNICODE).encode("utf-8"))
        emit(2, ("ö" * (CAP // 2 + 1) + UNICODE).encode("utf-8"))
        return 0
    if kind == "arguments":
        emit(1, json.dumps({"argv": arguments, "python": sys.executable,
             "encoding": sys.stdout.encoding, "utf8_env": os.environ.get("PYTHONUTF8"),
             "io_env": os.environ.get("PYTHONIOENCODING")}, ensure_ascii=False).encode("utf-8"))
        return 0
    if kind == "large-frame":
        emit(1, encoded_frame("L" * 100000 + UNICODE))
        return 1
    if kind == "large-frame-human-tail":
        emit(1, encoded_frame("L" * 100000 + UNICODE) + b"\n\nHuman last no-newline")
        return 1
    if kind == "oversized-frame":
        emit(1, encoded_frame("L" * 4096))
        return 1
    if kind == "duplicate-frames":
        emit(1, encoded_frame() + b"\n" + encoded_frame() + b"\n")
        return 1
    if kind == "no-frame":
        emit(1, b"Human diagnostic only")
        return 1
    if kind == "malformed-frame":
        emit(1, b"{broken authored JSON")
        return 1
    if kind in {"timeout-tree", "cancel-tree", "inherited-pipe", "detached-pipe"}:
        write_pid(arguments[0])
        flags = subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP if kind == "detached-pipe" else 0
        subprocess.Popen([sys.executable, "-I", "-B", __file__, "grandchild", arguments[1]],
                         stdout=sys.stdout, stderr=sys.stderr, creationflags=flags)
        for _ in range(200):
            if Path(arguments[1]).exists():
                break
            time.sleep(.005)
        if kind in {"inherited-pipe", "detached-pipe"}:
            return 0
        time.sleep(8)  # Finite even when used against defective historical code.
        return 0
    if kind == "grandchild":
        write_pid(arguments[0])
        time.sleep(8)
        return 0
    raise ValueError("Unknown authored native control: " + kind)


if __name__ == "__main__":
    raise SystemExit(main())
